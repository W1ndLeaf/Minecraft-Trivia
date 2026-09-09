"""
build.py - Trivia Challenge: Minecraft trivia with rewards and punishments as a plain vanilla DATA PACK.

    python tools/build.py               builds dist/TriviaChallenge-<version>.zip for every target and lints them
    python tools/build.py 26            one target only (26, 1.21 or 1.16)
    python tools/build.py --copy        also copies the 26 jar into INSTANCE_MODS (author convenience)

Questions come from questions.xlsx (sheet "Questions"). One code base, three targets:
  26    Minecraft 26.x / 1.21.11+   pop-up dialog UI (dialogs exist since 1.21.6), macro functions
  1.21  Minecraft 1.21 - 1.21.1     no dialogs -> clickable chat UI, macro functions
  1.16  Minecraft 1.16.5            clickable chat UI, no macros / random / return, fewer attributes
Vanilla has no storage that survives a new world, so the DEFAULTS below are what every fresh world starts with.
"""
import glob
import json
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from questions_xlsx import read_questions  # noqa: E402

XLSX = os.path.join(ROOT, "questions.xlsx")
BUILD = os.path.join(ROOT, "build")
DIST = os.path.join(ROOT, "dist")
VERSION = "2.0.0"
MODID = "triviachallenge"
NS = "tq"
NAME = "Trivia Challenge"
TAG = "[Trivia] "
INSTANCE_MODS = r"C:\Users\mateo\curseforge\minecraft\Instances\StudyTime\mods"

# ----------------------------------------------------------------------------- baked defaults (fresh world)
DEFAULT_DIFF = 1              # 1 = Normal, 2 = Hard
DEFAULT_INTERVAL_MIN = 5      # minutes between questions
INTERVAL_MAX_MIN = 30         # slider range 1..this
DEFAULT_SIDEBAR = 1           # 1 = score display on
AUTO_MENU = False             # True = open the menu once per world on first join; False = chat hint only
DISABLED_REWARDS = []         # reward keys that ship switched off, e.g. ["freeze"]
DISABLED_PUNISHMENTS = []     # punishment keys that ship switched off, e.g. ["wither"]
FREEZE_SECONDS = 10
BODY_W = 380
INTERVAL_PRESETS = [1, 2, 3, 5, 10, 15, 20, 30]   # the clickable choices in the chat UI

# ----------------------------------------------------------------------------- targets
ARMOR_SLOT = {"head": 103, "chest": 102, "legs": 101, "feet": 100}
TARGETS = {
    "26": dict(
        label="Minecraft 26.x (also 1.21.11+)", pack_format=107, mcmeta_extra={"min_format": 88, "max_format": 107},
        fn_dir="function", tag_dir="tags/function", loot_dir="loot_table", pred_dir="predicate",
        dialogs=True, macros=True, random_cmd=True, numberformat=True, warden=True, jar=True, mc=26,
        attr={"max_health": "minecraft:max_health", "scale": "minecraft:scale", "jump": "minecraft:jump_strength",
              "reach": "minecraft:block_interaction_range", "fall": "minecraft:fall_damage_multiplier"},
        armor_path=lambda part: f"equipment.{part}", count_key="count", extra_key="components", count_val="1",
        infinite="infinite", gamerule={"feedback": "send_command_feedback", "log": "log_admin_commands"},
        item_replace=lambda slot, item: f"item replace entity @s {slot} with {item}", potion_fn="set_potion",
        forbidden=["sendCommandFeedback", "logAdminCommands", "generic.", "player.block", "Inventory[{Slot:10",
                   "replaceitem", " 1000000 ", "Count:", "#minecraft:air", "if dimension"]),
    "1.21": dict(
        label="Minecraft 1.21 - 1.21.1", pack_format=48, mcmeta_extra={},
        fn_dir="function", tag_dir="tags/function", loot_dir="loot_table", pred_dir="predicate",
        dialogs=False, macros=True, random_cmd=True, numberformat=True, warden=True, jar=True, mc=21,
        attr={"max_health": "minecraft:generic.max_health", "scale": "minecraft:generic.scale",
              "jump": "minecraft:generic.jump_strength", "reach": "minecraft:player.block_interaction_range",
              "fall": "minecraft:generic.fall_damage_multiplier"},
        armor_path=lambda part: f"Inventory[{{Slot:{ARMOR_SLOT[part]}b}}]", count_key="count", extra_key="components",
        count_val="1", infinite="infinite", gamerule={"feedback": "sendCommandFeedback", "log": "logAdminCommands"},
        item_replace=lambda slot, item: f"item replace entity @s {slot} with {item}", potion_fn="set_potion",
        forbidden=["send_command_feedback", "log_admin_commands", "dialog ", "equipment.", "minecraft:scale",
                   "minecraft:jump_strength", "minecraft:block_interaction_range", "minecraft:fall_damage_multiplier",
                   "minecraft:max_health", "replaceitem", " 1000000 ", "Count:", "#minecraft:air", "if dimension"]),
    "1.16": dict(
        label="Minecraft 1.16.5", pack_format=6, mcmeta_extra={},
        fn_dir="functions", tag_dir="tags/functions", loot_dir="loot_tables", pred_dir="predicates",
        dialogs=False, macros=False, random_cmd=False, numberformat=False, warden=False, jar=False, mc=16,
        attr={"max_health": "minecraft:generic.max_health"},
        armor_path=lambda part: f"Inventory[{{Slot:{ARMOR_SLOT[part]}b}}]", count_key="Count", extra_key="tag",
        count_val="1b", infinite="1000000", gamerule={"feedback": "sendCommandFeedback", "log": "logAdminCommands"},
        item_replace=lambda slot, item: f"replaceitem entity @s {slot} {item}", potion_fn="set_nbt",
        forbidden=["send_command_feedback", "dialog ", "equipment.", "item replace", "random value", "return",
                   "numberformat", "display name", " infinite", "if dimension", "#minecraft:air", "$(", "minecraft:marker",
                   "components", "warden", "wind_charged", "if function"]),
}


# ----------------------------------------------------------------------------- questions
def sanitize(s):
    return "".join(c if 32 <= ord(c) < 127 else "?" for c in str(s)).strip()


def load_questions():
    qs = []
    for q in read_questions(XLSX):
        q = dict(q)
        q["q"] = sanitize(q["q"])
        q["options"] = [sanitize(o) for o in q["options"]]
        qs.append(q)
    qs.sort(key=lambda q: 0 if q["pool"] == "N" else 1)   # Normal block first, then Hard: each pool is one range
    for n, q in enumerate(qs, 1):
        q["n"] = n
    ranges = {}
    for q in qs:
        a, b = ranges.get(q["pool"], (q["n"], q["n"]))
        ranges[q["pool"]] = (min(a, q["n"]), max(b, q["n"]))
    for code, label in (("N", "Normal"), ("H", "Hard")):
        pool = [q for q in qs if q["pool"] == code]
        for k, q in enumerate(pool, 1):
            q["no"], q["total"], q["pool_label"] = k, len(pool), label
    return qs, ranges


# ----------------------------------------------------------------------------- text helpers
def comp(text, color=None, **kw):
    d = {"text": text}
    if color:
        d["color"] = color
    d.update(kw)
    return d


def jcomp(obj):
    return json.dumps(obj, ensure_ascii=True, separators=(",", ":"))


def click(text, color, command, hover=None, **kw):
    """A clickable chat piece (legacy click/hover event format, valid in 1.16 - 1.21.4)."""
    d = comp(text, color, **kw)
    d["clickEvent"] = {"action": "run_command", "value": "/" + command}
    if hover:
        d["hoverEvent"] = {"action": "show_text", "contents": comp(hover)}
    return d


def sub(text, color="white"):
    return f"title @s subtitle {jcomp(comp(text, color))}"


def bar(text, color="gray", target="@s"):
    return f"title {target} actionbar {jcomp(comp(text, color))}"


def msg(text, color="gray"):
    return "tellraw @s " + jcomp({"text": TAG, "color": "gold", "extra": [{"text": text, "color": color}]})


def tell(parts):
    return "tellraw @s " + jcomp(parts)


def score(name, objective, color=None):
    d = {"score": {"name": name, "objective": objective}}
    if color:
        d["color"] = color
    return d


# ----------------------------------------------------------------------------- rewards / punishments
# rewards: (key, short name, TITLE, subtitle, color, needs, commands)   needs: attribute keys / flags the target must have
def REWARDS(t):
    A = t["attr"]
    inf = t["infinite"]
    return [r for r in [
        ("tools", "Classic tools", "Classic tools", "water bucket + boat", "green", [], ["give @s minecraft:water_bucket", "give @s minecraft:oak_boat"]),
        ("nightvision", "Night Vision", "Night Vision", "permanent for this world", "aqua", [], ["tag @s add tq_e_nv", f"effect give @s minecraft:night_vision {inf} 0 true"]),
        ("slowfall", "Slow Falling", "Slow Falling", "permanent for this world", "aqua", [], ["tag @s add tq_e_sf", f"effect give @s minecraft:slow_falling {inf} 0 true"]),
        ("spawn", "Safe spawn", "Safe spawn", "your respawn point is now here", "green", [], ["spawnpoint @s ~ ~ ~"]),
        ("heart", "+1 heart", "+1 heart", "max health up, permanent", "gold", ["max_health"], [
            f"execute store result score #h tq_cfg run attribute @s {A.get('max_health')} base get",
            "scoreboard players add #h tq_cfg 2",
            "execute if score #h tq_cfg matches 41.. run scoreboard players set #h tq_cfg 40",
            "function tq:reward/heart_set"]),
        ("belly", "Full belly", "Full belly", "hunger and saturation restored", "green", [], ["effect give @s minecraft:saturation 1 10 true"]),
        ("regen", "Second wind", "Second wind", "Regeneration II for 15 seconds", "green", [], ["effect give @s minecraft:regeneration 15 1 true"]),
        ("cobble", "64 Cobblestone", "64 Cobblestone", "check your inventory", "green", [], ["give @s minecraft:cobblestone 64"]),
        ("shield", "Shield", "A Shield", "check your inventory", "green", [], ["give @s minecraft:shield"]),
        ("pickaxe", "Iron pickaxe", "Iron pickaxe", "check your inventory", "green", [], ["give @s minecraft:iron_pickaxe"]),
        ("freeze", "Time freeze", "TIME FREEZE", f"{FREEZE_SECONDS} s - every mob stops, you do not", "aqua", [], ["function tq:freeze/start"]),
        ("item", "Useful item", "Useful item", "check your inventory", "green", [], ["loot give @s loot tq:useful_item"]),
        ("bow", "Bow + arrows", "Bow + 8 arrows", "check your inventory", "green", [], ["give @s minecraft:bow", "give @s minecraft:arrow 8"]),
        ("fireres", "Fire Res 3 min", "Fire Resistance", "3 minutes", "gold", [], ["effect give @s minecraft:fire_resistance 180 0 true"]),
        ("haste", "Haste 3 min", "Haste I", "3 minutes", "gold", [], ["effect give @s minecraft:haste 180 0 true"]),
        ("speed", "Speed 3 min", "Speed I", "3 minutes", "gold", [], ["effect give @s minecraft:speed 180 0 true"]),
        ("reach", "Long arms", "Long arms", "reach +2 blocks for 3 minutes", "gold", ["reach"], [
            f"attribute @s {A.get('reach')} base set 6.5", "scoreboard players set @s tq_t_reach 180"]),
        ("tiny", "Tiny", "TINY", "half size for 3 minutes", "aqua", ["scale"], [
            f"attribute @s {A.get('scale')} base set 0.5", "scoreboard players set @s tq_t_scale 180"]),
    ] if all(n in A for n in r[5])]


# punishments: (key, short name, TITLE (<= 14 chars, it is the big text), detail, needs, commands)
def PUNISHMENTS(t):
    A = t["attr"]
    return [p for p in [
        ("grounded", "Grounded", "GROUNDED", "no jumping for 2 min", ["jump"], [f"attribute @s {A.get('jump')} base set 0", "scoreboard players set @s tq_t_jump 120"]),
        ("giant", "Giant", "GIANT", "double size for 3 min", ["scale"], [f"attribute @s {A.get('scale')} base set 2", "scoreboard players set @s tq_t_scale 180"]),
        ("shortarms", "Short arms", "SHORT ARMS", "reach 2 blocks for 3 min", ["reach"], [f"attribute @s {A.get('reach')} base set 2", "scoreboard players set @s tq_t_reach 180"]),
        ("glassbones", "Glass bones", "GLASS BONES", "fall damage x3 for 5 min", ["fall"], [f"attribute @s {A.get('fall')} base set 3", "scoreboard players set @s tq_t_fall 300"]),
        ("slow", "Slowness II", "SLOWNESS II", "2 min", [], ["effect give @s minecraft:slowness 120 1 true"]),
        ("fatigue", "Mining Fatigue", "MINING FATIGUE", "3 min", [], ["effect give @s minecraft:mining_fatigue 180 0 true"]),
        ("levitate", "Levitation V", "LEVITATION V", "8 seconds - going up, fast", [], ["effect give @s minecraft:levitation 8 4 true"]),
        ("lost", "Lost", "LOST", "300 blocks away", [], ["function tq:penalty/lost_go"]),
        ("mobs", "Mob wave", "MOB WAVE", "20 hostile mobs", [], ["scoreboard players set #i tq_cfg 0", "function tq:penalty/mob_loop"]),
        ("warden", "Warden", "A WARDEN", "good luck", ["warden"], ["playsound minecraft:entity.warden.emerge master @s ~ ~ ~ 1 1", "summon minecraft:warden ~ ~ ~"]),
        ("skydrop", "Sky drop", "SKY DROP", "200 blocks up", [], ["tp @s ~ ~200 ~"]),
        ("wither", "Wither", "THE WITHER", "it is right behind you", [], ["playsound minecraft:entity.wither.spawn master @s ~ ~ ~ 1 1", "summon minecraft:wither ~3 ~ ~3"]),
        ("starve", "Starvation", "STARVATION", "hunger drained + Hunger III 3 min", [], ["effect give @s minecraft:hunger 3 254 true", "function tq:penalty/starve2"]),
        ("blind", "Lights out", "LIGHTS OUT", "Blindness 1 min (no sprinting)", [], ["effect give @s minecraft:blindness 60 0 true"]),
        ("butterfingers", "Butterfingers", "BUTTERFINGERS", "hotbar dropped", [], ["function tq:penalty/drop_hotbar"]),
        ("spawn", "Back to spawn", "BACK TO SPAWN", "where you first joined", [], ["function tq:penalty/to_spawn"]),
        ("poison", "Poison II", "POISON II", "30 seconds", [], ["effect give @s minecraft:poison 30 1 true"]),
        ("feeble", "Feeble", "FEEBLE", "Weakness II for 3 min", [], ["effect give @s minecraft:weakness 180 1 true"]),
        ("phantoms", "Phantom night", "PHANTOM NIGHT", "night + 6 phantoms", [], ["time set night", "function tq:penalty/phantoms"]),
        ("creepers", "Creeper ambush", "CREEPER AMBUSH", "4 charged creepers", [], ["scoreboard players set #i tq_cfg 0", "function tq:penalty/creeper_loop"]),
        ("armorstrip", "Armor strip", "ARMOR STRIP", "armor dropped", [], ["function tq:penalty/drop_armor"]),
        ("toolbreak", "Tool break", "TOOL BREAK", "main hand destroyed", [], ["playsound minecraft:entity.item.break master @s ~ ~ ~ 1 1", t["item_replace"]("weapon.mainhand", "minecraft:air")]),
        ("buried", "Buried alive", "BURIED ALIVE", "15 blocks down", [], ["function tq:penalty/buried_go"]),
    ] if all((n in A) or (n == "warden" and t["warden"]) for n in p[4])]


# useful items, all with the same chance: (item id[:variant], count).  ITEM_MIN = first Minecraft minor that has it.
POOL_ITEMS = [
    ("iron_pickaxe", 1), ("iron_axe", 1), ("iron_shovel", 1), ("iron_sword", 1), ("iron_hoe", 1),
    ("bow", 1), ("crossbow", 1), ("fishing_rod", 1), ("shears", 1), ("mace", 1), ("elytra", 1),
    ("leather_helmet", 1), ("leather_chestplate", 1), ("chainmail_chestplate", 1), ("iron_helmet", 1), ("iron_boots", 1), ("iron_leggings", 1),
    ("bread", 8), ("cooked_beef", 6), ("cooked_porkchop", 6), ("cooked_chicken", 6), ("baked_potato", 8), ("cooked_cod", 6),
    ("golden_carrot", 4), ("golden_apple", 1), ("apple", 8), ("cooked_mutton", 6), ("pumpkin_pie", 3), ("cake", 1),
    ("cobblestone", 64), ("dirt", 32), ("oak_log", 16), ("oak_planks", 16), ("gravel", 32), ("sand", 32), ("stone_bricks", 16),
    ("glass", 8), ("ladder", 8), ("scaffolding", 12), ("torch", 16), ("hay_block", 4),
    ("water_bucket", 1), ("bucket", 1), ("lava_bucket", 1), ("oak_boat", 1), ("red_bed", 1),
    ("name_tag", 1), ("saddle", 1), ("firework_rocket", 4), ("flint_and_steel", 1),
    ("string", 8), ("shield", 1),
    ("iron_ingot", 4), ("coal", 8), ("gold_ingot", 4), ("flint", 8), ("leather", 4), ("stick", 8), ("bone", 4), ("bone_meal", 8),
    ("slime_ball", 4), ("redstone", 8), ("lapis_lazuli", 4), ("diamond", 1), ("emerald", 1), ("copper_ingot", 4),
    ("ender_pearl", 1), ("blaze_rod", 1), ("obsidian", 4), ("blaze_powder", 1), ("gunpowder", 2), ("nether_wart", 4),
    ("ghast_tear", 1), ("magma_cream", 1), ("quartz", 2), ("ender_eye", 1),
    ("potion:fire_resistance", 1), ("potion:healing", 1), ("potion:swiftness", 1), ("splash_potion:healing", 1),
    ("milk_bucket", 1), ("arrow", 16), ("tnt", 2), ("suspicious_stew:regeneration", 1),
    ("honey_bottle", 1), ("spyglass", 1), ("paper", 8), ("book", 1), ("snowball", 4), ("carved_pumpkin", 1),
]
ITEM_MIN = {"mace": 21, "copper_ingot": 17, "spyglass": 17}
BLOCK_ITEMS = {"cobblestone", "dirt", "oak_log", "oak_planks", "gravel", "sand", "stone_bricks", "glass", "ladder",
               "scaffolding", "torch", "hay_block", "tnt", "obsidian", "carved_pumpkin", "red_bed", "cake", "nether_wart"}


def useful_loot_table(t, known):
    entries, unknown = [], []
    for spec, count in POOL_ITEMS:
        base, arg = (spec.split(":", 1) + [None])[:2]
        if ITEM_MIN.get(base, 0) > t["mc"]:
            continue
        if base not in known:
            unknown.append(base)
        e = {"type": "minecraft:item", "name": "minecraft:" + base, "weight": 1, "functions": []}
        if count > 1:
            e["functions"].append({"function": "minecraft:set_count", "count": count})
        if base in ("potion", "splash_potion"):
            if t["potion_fn"] == "set_potion":
                e["functions"].append({"function": "minecraft:set_potion", "id": "minecraft:" + arg})
            else:
                e["functions"].append({"function": "minecraft:set_nbt", "tag": '{Potion:"minecraft:' + arg + '"}'})
        elif base == "suspicious_stew":
            e["functions"].append({"function": "minecraft:set_stew_effect", "effects": [{"type": "minecraft:" + arg, "duration": 7}]})
        if not e["functions"]:
            del e["functions"]
        entries.append(e)
    return {"type": "minecraft:generic", "pools": [{"rolls": 1, "entries": entries}]}, unknown

# ----------------------------------------------------------------------------- 26: dialog UI
def btn(label, cmd, color="white", width=100, tooltip=None):
    b = {"label": comp(label, color), "width": width,
         "action": {"type": "minecraft:run_command", "command": cmd}}
    if tooltip:
        b["tooltip"] = comp(tooltip)
    return b


def question_dialog(q):
    extra = [comp(q["q"], "white"), comp("\n\n")]
    for i, o in enumerate(q["options"]):
        extra += [comp("ABCD"[i] + ")  ", "yellow", bold=True), comp(o, "white"), comp("\n")]
    actions = [btn("ABCD"[i], f"trigger tq_answer set {i + 1}", "yellow", 70, q["options"][i])
               for i in range(len(q["options"]))]
    return {
        "type": "minecraft:multi_action",
        "title": comp(f"{NAME}  -  {q['pool_label']} {q['no']}/{q['total']}", "gold"),
        "external_title": comp(f"{NAME} question"),
        "pause": True,
        "can_close_with_escape": True,
        "after_action": "close",
        "body": [{"type": "minecraft:plain_message", "contents": {"text": "", "extra": extra}, "width": BODY_W}],
        "columns": len(actions),
        "actions": actions,
        "exit_action": {"label": comp("Skip (counts as wrong)", "gray"), "width": 150,
                        "action": {"type": "minecraft:run_command", "command": "trigger tq_answer set 9"}},
    }


def body_lines(lines):
    return [{"type": "minecraft:plain_message", "contents": comp(t, c), "width": BODY_W} for t, c in lines]


# codes for the dialog-only trigger  /trigger tq_ui set <code>
UI_STATS, UI_CLOSE, UI_DIFF, UI_STATS_OPT = 3, 4, 5, 6       # STATS re-shows the menu, STATS_OPT the options
UI_ALL_R, UI_ALL_P, UI_ALL_P_OFF, UI_ALL_R_OFF = 11, 12, 13, 14
UI_R_BASE, UI_P_BASE = 100, 200                               # 101.. toggle reward N, 201.. toggle punishment N
ALL_ON, ALL_OFF = 100, 101                                    # /trigger tq_tr|tq_tp set 100|101 (chat lists)


def ui_dialog(title, lines, actions, columns, inputs=None):
    """A settings screen, shown INLINE from a macro function (dialog show @s {...}) so labels and colors carry
    the live state ($(...) placeholders). Vanilla lays a dialog out top to bottom: body, inputs, button grid
    (an incomplete last row is centred), footer exit button - so Close is always the footer button and a button
    can never sit beside an input. With after_action "none" not even the exit button closes by itself, so Close
    runs the close trigger and the tick function answers with `dialog clear`; ESC closes directly."""
    d = {
        "type": "minecraft:multi_action",
        "title": comp(title, "gold"),
        "external_title": comp(title),
        "pause": False,
        "can_close_with_escape": True,
        "after_action": "none",
        "body": body_lines(lines),
        "columns": columns,
        "actions": actions,
        "exit_action": btn("Close", f"trigger tq_ui set {UI_CLOSE}", "white", 110),
    }
    if inputs:
        d["inputs"] = inputs
    return d


def ui_macro(d):
    """One macro line: $dialog show @s <inline SNBT>. "@@iv@@" marks a number that must stay unquoted."""
    s = json.dumps(d, ensure_ascii=True, separators=(",", ":")).replace('"@@iv@@"', "$(iv)")
    assert "\\n" not in s and "\\u" not in s, "inline dialog text must be plain ASCII single lines"
    return "$dialog show @s " + s


def menu_dialog(counts):
    return ui_dialog(NAME, [
        (NAME.upper(), "gold"),
        (f"{counts['N']} Normal and {counts['H']} Hard Minecraft trivia questions. Every few minutes one pops up: "
         "answer A-D, or ESC to skip (= wrong).", "white"),
        ("Correct = a random reward. Wrong = a random punishment. Whatever hits you becomes rarer next time, "
         "so everything comes around.", "gray"),
        ("Now: $(dn) difficulty, a question every $(iv) min.", "yellow"),
        ("Chat: /trigger tq_menu = this menu, /trigger tq_ask = a question now, /trigger tq_opts = options.", "gray"),
    ], [
        btn("Ask me now", "trigger tq_ask", "green", 130),
        btn("Options", "trigger tq_opts", "aqua", 130),
        btn("Score display: $(sd)", f"trigger tq_ui set {UI_STATS}", "$(sdc)", 130),
    ], 3)


def options_dialog():
    slider = {"type": "minecraft:number_range", "key": "iv", "label": comp("Interval (minutes)"),
              "start": 1, "end": INTERVAL_MAX_MIN, "step": 1, "initial": "@@iv@@", "width": 200}
    # the template needs a literal $(iv); it arrives through storage so the macro function leaves it alone
    apply_iv = {"label": comp("Apply interval", "yellow"), "width": 200,
                "action": {"type": "minecraft:dynamic/run_command", "template": "$(tpl)"}}
    return ui_dialog(f"{NAME} - options", [
        ("OPTIONS", "gold"),
        ("Interval: now a question every $(iv) min. Move the slider, then press Apply interval.", "white"),
        ("Difficulty: Normal = hard for most players, Hard = deep-cut Minecraft knowledge. Click to switch.", "gray"),
        ("Rewards... / Punishments...: every item can be switched off (red) or on (green).", "gray"),
    ], [
        apply_iv,
        btn("Difficulty: $(dn)", f"trigger tq_ui set {UI_DIFF}", "$(dc)", 200),
        btn("Score display: $(sd)", f"trigger tq_ui set {UI_STATS_OPT}", "$(sdc)", 200),
        btn("Rewards...", "trigger tq_rl", "aqua", 200),
        btn("Punishments...", "trigger tq_pl", "aqua", 200),
        btn("< Menu", "trigger tq_menu", "gray", 200),
    ], 1, inputs=[slider])


def toggle_dialog(kind, items, base, all_on, all_off):
    actions = [btn(name, f"trigger tq_ui set {base + i}", f"$(c{i})", 110) for i, name in enumerate(items, 1)]
    actions += [btn("All on", f"trigger tq_ui set {all_on}", "aqua", 110),
                btn("All off", f"trigger tq_ui set {all_off}", "aqua", 110),
                btn("< Options", "trigger tq_opts", "gray", 110)]
    return ui_dialog(f"{NAME} - {kind}", [
        (kind.upper(), "gold"),
        ("Green = on, red = off. Click an item to switch it; off items are never rolled.", "white"),
        ("All off = none of these can happen.", "gray"),
    ], actions, 4)


# ----------------------------------------------------------------------------- 1.21 / 1.16: clickable chat UI
def chat_header(text):
    return tell({"text": f"---------- {text} ----------", "color": "gold", "bold": True})


def chat_question(q):
    lines = [chat_header(f"TRIVIA - {q['pool_label']} {q['no']}/{q['total']}"), tell(comp(q["q"], "white"))]
    for i, o in enumerate(q["options"]):
        letter, cmd = "ABCD"[i], f"trigger tq_answer set {i + 1}"
        lines.append(tell([click(f"[{letter}] ", "yellow", cmd, f"Answer {letter}", bold=True),
                           click(o, "white", cmd, f"Answer {letter}")]))
    lines.append(tell([click("[Skip]", "dark_gray", "trigger tq_answer set 9", "Skip = counts as wrong"),
                       comp(" (counts as wrong)", "dark_gray")]))
    return lines


def chat_menu(counts):
    def buttons(sd_label, sd_color):
        return tell([click("[Ask me now]", "green", "trigger tq_ask", "A question right now"), comp("   "),
                     click("[Options]", "aqua", "trigger tq_opts", "Difficulty, interval, rewards, punishments"),
                     comp("   "), click(sd_label, sd_color, "trigger tq_stats", "Show / hide the sidebar")])

    def now(diff_text, diff_color):
        return tell([comp("Now: ", "yellow"), comp(diff_text, diff_color), comp(" difficulty, a question every ", "yellow"),
                     score("#interval_min", "tq_cfg", "aqua"), comp(" min.  Correct ", "yellow"), score("@s", "tq_right", "green"),
                     comp(" / wrong ", "yellow"), score("@s", "tq_wrong", "red")])
    return [
        chat_header("TRIVIA CHALLENGE"),
        tell(comp(f"{counts['N']} Normal and {counts['H']} Hard Minecraft trivia questions. Every few minutes one appears "
                  "in chat: click A-D to answer, or Skip (= wrong). Correct = a random reward, wrong = a random "
                  "punishment; whatever hits you becomes rarer next time.", "gray")),
        "execute if score #diff tq_cfg matches 2 run " + now("HARD", "red"),
        "execute unless score #diff tq_cfg matches 2 run " + now("NORMAL", "green"),
        "execute if score #sidebar tq_cfg matches 1 run " + buttons("[Score display: ON]", "green"),
        "execute unless score #sidebar tq_cfg matches 1 run " + buttons("[Score display: OFF]", "red"),
    ]


def chat_options():
    def diff(cur):
        return tell([comp("Difficulty: ", "white"),
                     click("[NORMAL]", "green" if cur == 1 else "gray", "trigger tq_diff set 1", "Normal: hard for most players"),
                     comp("  "), click("[HARD]", "red" if cur == 2 else "gray", "trigger tq_diff set 2", "Hard: deep-cut knowledge")])
    interval = [comp("Interval: now every ", "white"), score("#interval_min", "tq_cfg", "aqua"), comp(" min. Set: ", "white")]
    for n in INTERVAL_PRESETS:
        interval += [click(f"[{n}]", "aqua", f"trigger tq_interval set {n}", f"A question every {n} min"), comp(" ")]

    def rest(sd_label, sd_color):
        return tell([click(sd_label, sd_color, "trigger tq_stats", "Show / hide the sidebar"), comp("  "),
                     click("[Rewards...]", "aqua", "trigger tq_rl", "Switch rewards on / off"), comp("  "),
                     click("[Punishments...]", "aqua", "trigger tq_pl", "Switch punishments on / off"), comp("  "),
                     click("[Menu]", "gray", "trigger tq_menu", "Back to the menu")])
    return [
        chat_header("OPTIONS"),
        "execute if score #diff tq_cfg matches 2 run " + diff(2),
        "execute unless score #diff tq_cfg matches 2 run " + diff(1),
        tell(interval),
        "execute if score #sidebar tq_cfg matches 1 run " + rest("[Score display: ON]", "green"),
        "execute unless score #sidebar tq_cfg matches 1 run " + rest("[Score display: OFF]", "red"),
    ]


def chat_list(kind, prefix, names, trig, reopen):
    lines = [chat_header(f"{kind.upper()} - click to switch on / off")]
    for i, name in enumerate(names, 1):
        lines.append(f"execute if score {prefix}{i} tq_on matches 1 run "
                     + tell(click(f"[ON]  {name}", "green", f"trigger {trig} set {i}", "Click to switch OFF")))
        lines.append(f"execute unless score {prefix}{i} tq_on matches 1 run "
                     + tell(click(f"[OFF] {name}", "red", f"trigger {trig} set {i}", "Click to switch ON")))
    lines.append(tell([click("[All on]", "aqua", f"trigger {trig} set {ALL_ON}"), comp("  "),
                       click("[All off]", "aqua", f"trigger {trig} set {ALL_OFF}", "None of these can happen"), comp("  "),
                       click("[Show list again]", "gray", f"trigger {reopen}"), comp("  "),
                       click("[Options]", "gray", "trigger tq_opts")]))
    return lines

# ----------------------------------------------------------------------------- functions: core game
def functions(t, qs, ranges):
    fn = {}
    A = t["attr"]
    nmin, nmax = ranges["N"]
    hmin, hmax = ranges["H"]
    ncount, hcount = nmax - nmin + 1, hmax - hmin + 1
    counts = {"N": ncount, "H": hcount}
    rewards, punishments = REWARDS(t), PUNISHMENTS(t)
    rkeys, rnames = [r[0] for r in rewards], [r[1] for r in rewards]
    pkeys, pnames = [p[0] for p in punishments], [p[1] for p in punishments]
    triggers = ["tq_ask", "tq_interval", "tq_menu", "tq_opts", "tq_diff", "tq_stats", "tq_rl", "tq_pl", "tq_tr", "tq_tp",
                "tq_status"] + (["tq_ui", "tq_iv"] if t["dialogs"] else [])
    timers = [(k, o, A[k], d) for k, o, d in (("jump", "tq_t_jump", "0.42"), ("scale", "tq_t_scale", "1"),
                                              ("reach", "tq_t_reach", "4.5"), ("fall", "tq_t_fall", "1")) if k in A]

    def reset_trigger(x):
        return [f"scoreboard players set @s {x} 0", f"scoreboard players enable @s {x}"]

    # ---------------- load
    fn["load"] = [f"scoreboard objectives add {o} dummy" for o in (
        "tq_timer", "tq_state", "tq_qid", "tq_last", "tq_wait", "tq_delay", "tq_tmp", "tq_right", "tq_wrong", "tq_cfg",
        "tq_key", "tq_seen", "tq_w", "tq_on", "tq_t_jump", "tq_t_scale", "tq_t_reach", "tq_t_fall")] + [
        "scoreboard objectives add tq_answer trigger",
    ] + [f"scoreboard objectives add {x} trigger" for x in triggers] + [
        "scoreboard objectives add tq_health health",
        "scoreboard objectives add tq_deaths deathCount",
        "scoreboard objectives add tq_leave minecraft.custom:minecraft.leave_game",
        "scoreboard objectives add tq_score dummy " + jcomp(comp(NAME, "gold")),
    ]
    if t["numberformat"]:   # fixed order: hidden scores 3/2/1, the values live in the display names
        fn["load"] += [
            "scoreboard objectives modify tq_score numberformat blank",
            "scoreboard players set Nxt tq_score 3", "scoreboard players set Cor tq_score 2", "scoreboard players set Wrg tq_score 1",
            'scoreboard players display name Nxt tq_score {"text":"Next question: - min","color":"aqua"}',
            'scoreboard players display name Cor tq_score {"text":"Correct: -","color":"green"}',
            'scoreboard players display name Wrg tq_score {"text":"Wrong: -","color":"red"}',
        ]
    else:                   # 1.16: numbers cannot be hidden, so the values ARE the scores (classic look)
        fn["load"] += [
            "team add tq_nxt", "team add tq_cor", "team add tq_wrg",
            "team modify tq_nxt prefix " + jcomp(comp("Next question: ", "aqua")), "team modify tq_nxt color aqua",
            "team modify tq_cor color green", "team modify tq_wrg color red",
            "team join tq_nxt minutes", "team join tq_cor Correct", "team join tq_wrg Wrong",
            "scoreboard players set minutes tq_score 0", "scoreboard players set Correct tq_score 0", "scoreboard players set Wrong tq_score 0",
        ]
    fn["load"] += [
        f"execute unless score #sidebar tq_cfg matches 0..1 run scoreboard players set #sidebar tq_cfg {DEFAULT_SIDEBAR}",
        "execute if score #sidebar tq_cfg matches 1 run scoreboard objectives setdisplay sidebar tq_score",
        f"execute unless score #interval_min tq_cfg matches 1.. run scoreboard players set #interval_min tq_cfg {DEFAULT_INTERVAL_MIN}",
        f"execute unless score #interval tq_cfg matches 1.. run scoreboard players set #interval tq_cfg {DEFAULT_INTERVAL_MIN * 1200}",
        f"execute unless score #diff tq_cfg matches 1..2 run scoreboard players set #diff tq_cfg {DEFAULT_DIFF}",
        "scoreboard players set #1200 tq_cfg 1200", "scoreboard players set #two tq_cfg 2", "scoreboard players set #five tq_cfg 5",
        "scoreboard players set #eight tq_cfg 8", "scoreboard players set #twelve tq_cfg 12", "scoreboard players set #clock tq_cfg 0",
        f"scoreboard players set #N_min tq_cfg {nmin}", f"scoreboard players set #N_max tq_cfg {nmax}",
        f"scoreboard players set #H_min tq_cfg {hmin}", f"scoreboard players set #H_max tq_cfg {hmax}",
        f"execute unless score #unseen_N tq_cfg matches 0..{ncount} run scoreboard players set #unseen_N tq_cfg {ncount}",
        f"execute unless score #unseen_H tq_cfg matches 0..{hcount} run scoreboard players set #unseen_H tq_cfg {hcount}",
    ]
    for i, key in enumerate(rkeys, 1):
        fn["load"].append(f"execute unless score r{i} tq_w matches 1.. run scoreboard players set r{i} tq_w 64")
        fn["load"].append(f"execute unless score r{i} tq_on matches 0..1 run scoreboard players set r{i} tq_on {0 if key in DISABLED_REWARDS else 1}")
    for i, key in enumerate(pkeys, 1):
        fn["load"].append(f"execute unless score p{i} tq_w matches 1.. run scoreboard players set p{i} tq_w 64")
        fn["load"].append(f"execute unless score p{i} tq_on matches 0..1 run scoreboard players set p{i} tq_on {0 if key in DISABLED_PUNISHMENTS else 1}")
    fn["load"] += [
        # the "Triggered [tq_x]" lines are /trigger command feedback: switch it off once (admin/feedback_on restores).
        # Own function on purpose: if a command in it ever fails to parse, load itself still runs.
        "execute unless score #fb_set tq_cfg matches 1.. run function tq:feedback_off",
        "function tq:freeze/stop_silent",
        "function tq:keys",
        "function tq:enable_triggers",
        "tellraw @a " + jcomp({"text": TAG, "color": "gold", "extra": [
            {"text": f"loaded {len(qs)} questions. ", "color": "gray"},
            {"text": "/trigger tq_menu", "color": "yellow"}, {"text": " = menu, ", "color": "gray"},
            {"text": "/trigger tq_ask", "color": "yellow"}, {"text": " = a question now.", "color": "gray"}]}),
    ]
    G = t["gamerule"]
    fn["feedback_off"] = [f"gamerule {G['feedback']} false", f"gamerule {G['log']} false", "scoreboard players set #fb_set tq_cfg 1"]
    fn["keys"] = [f"scoreboard players set q{q['n']} tq_key {q['correct'] + 1}" for q in qs]

    def answers(q):   # ans = subtitle text (kept short so it fits at any GUI scale), ansfull = whole answer for chat
        full = "ABCD"[q["correct"]] + " - " + q["options"][q["correct"]]
        short = full if len(full) <= 36 else full[:33].rstrip() + "..."
        return f"execute if score @s tq_qid matches {q['n']} run data merge storage tq:tmp " + \
            "{ans:" + json.dumps(short, ensure_ascii=True) + ",ansfull:" + json.dumps(full, ensure_ascii=True) + "}"
    fn["ans"] = [answers(q) for q in qs]
    # tq_answer stays enabled too: a click on an old answer link is then silently swallowed by the tick cleanup
    fn["enable_triggers"] = [f"scoreboard players enable @a {x}" for x in triggers + ["tq_answer"]]

    # ---------------- tick
    fn["tick"] = [
        "scoreboard players add #clock tq_cfg 1",
        "execute as @a[scores={tq_leave=1..}] run function tq:on_join",
        "execute as @a[tag=!tq_init] at @s run function tq:first_join",
        "execute as @a[scores={tq_deaths=1..}] run function tq:on_death",
        "execute as @a[scores={tq_delay=1..}] run function tq:delay_tick",
    ] + [f"execute as @a[scores={{{x}=1..}}] at @s run function tq:trig/{x[3:]}" for x in triggers] + [
        "execute as @a[scores={tq_state=0,tq_answer=1..}] run scoreboard players set @s tq_answer 0",
        "execute as @a[scores={tq_state=1,tq_answer=1..}] at @s run function tq:answered",
        "execute as @a[scores={tq_state=0,tq_health=1..},gamemode=!spectator] at @s run function tq:timer_tick",
        "execute as @a[scores={tq_state=1..}] run function tq:wait_tick",
        "execute if score #clock tq_cfg matches 20.. run function tq:second",
    ]
    fn["second"] = [
        "scoreboard players set #clock tq_cfg 0",
        "function tq:enable_triggers",
        f"execute as @a[tag=tq_e_nv] run effect give @s minecraft:night_vision {t['infinite']} 0 true",
        f"execute as @a[tag=tq_e_sf] run effect give @s minecraft:slow_falling {t['infinite']} 0 true",
    ] + [f"execute as @a[scores={{{o}=1..}}] run function tq:timers/{k}" for k, o, _, _ in timers] + [
        "execute if score #freeze tq_cfg matches 1.. run function tq:freeze/tick",
        "execute if score #sidebar tq_cfg matches 1 as @a[limit=1] run function tq:sidebar",
    ]
    for k, o, attr, default in timers:
        fn[f"timers/{k}"] = [
            f"scoreboard players remove @s {o} 1",
            f"execute if score @s {o} matches ..0 run attribute @s {attr} base set {default}",
            f"execute if score @s {o} matches ..0 run scoreboard players reset @s {o}",
        ]
    fn["timers/clear_all"] = [f"attribute @s {attr} base set {default}" for _, _, attr, default in timers] + \
        [f"scoreboard players reset @s {o}" for _, o, _, _ in timers] + ["scoreboard players reset @s tq_tmp"]
    spawn_marker = ("summon minecraft:area_effect_cloud ~ ~ ~ {Tags:[\"tq_spawn\"],Duration:2147483647,Age:-2147483648,"
                    "WaitTime:-2147483648,Radius:0.0f}")
    fn["first_join"] = [
        "tag @s add tq_init",
        "scoreboard players set @s tq_state 0", "scoreboard players set @s tq_timer 0",
        "scoreboard players set @s tq_wait 0", "scoreboard players set @s tq_leave 0",
        "execute unless score @s tq_right matches 0.. run scoreboard players set @s tq_right 0",
        "execute unless score @s tq_wrong matches 0.. run scoreboard players set @s tq_wrong 0",
    ] + ([
        "execute unless data storage tq:spawn x run data modify storage tq:spawn x set from entity @s Pos[0]",
        "execute unless data storage tq:spawn y run data modify storage tq:spawn y set from entity @s Pos[1]",
        "execute unless data storage tq:spawn z run data modify storage tq:spawn z set from entity @s Pos[2]",
    ] if t["macros"] else [
        "execute unless score #spawn_set tq_cfg matches 1 run " + spawn_marker,
        "scoreboard players set #spawn_set tq_cfg 1",
    ]) + (["scoreboard players set @s tq_delay 80"] if AUTO_MENU else []) + [
        "tellraw @s " + jcomp({"text": TAG, "color": "gold", "extra": [
            {"text": "Welcome! ", "color": "gray"}, {"text": "/trigger tq_menu", "color": "yellow"},
            {"text": " = menu and options, ", "color": "gray"}, {"text": "/trigger tq_ask", "color": "yellow"},
            {"text": " = a question now.", "color": "gray"}]}),
    ]
    fn["delay_tick"] = ["scoreboard players remove @s tq_delay 1",
                        "execute if score @s tq_delay matches 0 run function tq:ui/menu_show"]
    fn["on_death"] = ["scoreboard players set @s tq_deaths 0", "function tq:timers/clear_all"]
    fn["on_join"] = ["scoreboard players set @s tq_leave 0",
                     "execute if score @s tq_state matches 1.. run scoreboard players set @s tq_wait 2340"]
    fn["timer_tick"] = ["scoreboard players add @s tq_timer 1",
                        "execute if score @s tq_timer >= #interval tq_cfg run function tq:ask"]

    # ---------------- asking: exact "every question once, then reset" picking, no macros needed
    fn["ask"] = [
        "scoreboard players set @s tq_timer 0",
        "function tq:pick",
        "scoreboard players set @s tq_state 1", "scoreboard players set @s tq_wait 0",
        "scoreboard players reset @s tq_answer", "scoreboard players enable @s tq_answer",
        "function tq:show_q",
    ]
    if t["dialogs"]:
        fn["show_q"] = [f"execute if score @s tq_qid matches {q['n']} run dialog show @s tq:q_{q['n']}" for q in qs]
    else:
        fn["show_q"] = ["playsound minecraft:block.note_block.pling master @s ~ ~ ~ 1 1.5",
                        bar("A trivia question is waiting in chat - open the chat and click an answer", "gold")] + \
            [f"execute if score @s tq_qid matches {q['n']} run function tq:q/{q['n']}" for q in qs]
        for q in qs:
            fn[f"q/{q['n']}"] = chat_question(q)
    if t["random_cmd"]:
        fn["rng"] = ["execute store result score #rnd tq_cfg run random value 0..2147483647"]
    else:   # 1.16: the first int of a fresh entity UUID is a random 32-bit number
        fn["rng"] = ['execute at @s run summon minecraft:area_effect_cloud ~ ~ ~ {Tags:["tq_rng"],Duration:1}',
                     "execute store result score #rnd tq_cfg run data get entity @e[type=minecraft:area_effect_cloud,tag=tq_rng,limit=1] UUID[0]",
                     "kill @e[type=minecraft:area_effect_cloud,tag=tq_rng]"]
    fn["pick"] = [
        "execute if score #diff tq_cfg matches 1 if score #unseen_N tq_cfg matches ..0 run function tq:reset_n",
        "execute if score #diff tq_cfg matches 2 if score #unseen_H tq_cfg matches ..0 run function tq:reset_h",
        "execute if score #diff tq_cfg matches 1 run scoreboard players operation #unseen tq_cfg = #unseen_N tq_cfg",
        "execute if score #diff tq_cfg matches 2 run scoreboard players operation #unseen tq_cfg = #unseen_H tq_cfg",
        "function tq:rng",
        "scoreboard players operation #k tq_cfg = #rnd tq_cfg",
        "scoreboard players operation #k tq_cfg %= #unseen tq_cfg",   # %= is floor-mod: never negative
        "scoreboard players add #k tq_cfg 1",
        "scoreboard players set #done tq_cfg 0",
        "execute if score #diff tq_cfg matches 1 run function tq:pick_walk_n",
        "execute if score #diff tq_cfg matches 2 run function tq:pick_walk_h",
        "execute if score #done tq_cfg matches 0 if score #diff tq_cfg matches 1 run scoreboard players operation @s tq_qid = #N_min tq_cfg",
        "execute if score #done tq_cfg matches 0 if score #diff tq_cfg matches 2 run scoreboard players operation @s tq_qid = #H_min tq_cfg",
        "function tq:pick_mark",
        "execute if score #diff tq_cfg matches 1 run scoreboard players remove #unseen_N tq_cfg 1",
        "execute if score #diff tq_cfg matches 2 run scoreboard players remove #unseen_H tq_cfg 1",
    ]
    for name, lo, hi in (("pick_walk_n", nmin, nmax), ("pick_walk_h", hmin, hmax)):
        walk = []
        for n in range(lo, hi + 1):   # the k-th unseen question of the pool wins
            walk.append(f"execute if score #done tq_cfg matches 0 unless score q{n} tq_seen matches 1 run scoreboard players remove #k tq_cfg 1")
            walk.append(f"execute if score #done tq_cfg matches 0 if score #k tq_cfg matches ..0 run scoreboard players set @s tq_qid {n}")
            walk.append(f"execute if score #done tq_cfg matches 0 if score #k tq_cfg matches ..0 run scoreboard players set #done tq_cfg 1")
        fn[name] = walk
    fn["pick_mark"] = [f"execute if score @s tq_qid matches {q['n']} run scoreboard players set q{q['n']} tq_seen 1" for q in qs]
    fn["reset_n"] = [f"scoreboard players reset q{n} tq_seen" for n in range(nmin, nmax + 1)] + [
        f"scoreboard players set #unseen_N tq_cfg {ncount}", msg("every Normal question has been asked - starting the pool over.", "yellow")]
    fn["reset_h"] = [f"scoreboard players reset q{n} tq_seen" for n in range(hmin, hmax + 1)] + [
        f"scoreboard players set #unseen_H tq_cfg {hcount}", msg("every Hard question has been asked - starting the pool over.", "yellow")]

    # ---------------- answering
    fn["answered"] = [
        "scoreboard players operation @s tq_tmp = @s tq_answer",
        "scoreboard players reset @s tq_answer",
        "scoreboard players set @s tq_state 0", "scoreboard players set @s tq_timer 0", "scoreboard players set @s tq_wait 0",
        "function tq:judge",
        "execute if score #sidebar tq_cfg matches 1 run function tq:sidebar",
    ]
    fn["judge"] = [f"execute if score @s tq_qid matches {q['n']} run scoreboard players set #key tq_cfg {q['correct'] + 1}" for q in qs] + [
        "execute if score @s tq_tmp = #key tq_cfg run function tq:right",
        "execute unless score @s tq_tmp = #key tq_cfg run function tq:wrong",
    ]
    fn["right"] = [
        "scoreboard players set @s tq_last 1", "scoreboard players add @s tq_right 1",
        "playsound minecraft:entity.player.levelup master @s ~ ~ ~ 0.8 1.2",
        "function tq:reward/random",
    ]
    fn["wrong"] = [
        "scoreboard players set @s tq_last 2", "scoreboard players add @s tq_wrong 1",
        "playsound minecraft:entity.villager.no master @s ~ ~ ~ 1 0.8",
        "function tq:ans",
        'title @s subtitle {"text":"Correct: ","color":"yellow","extra":[{"nbt":"ans","storage":"tq:tmp","color":"white"}]}',
        "function tq:penalty/random",
    ]
    fn["wait_tick"] = ["scoreboard players add @s tq_wait 1",
                       "execute if score @s tq_wait matches 2400.. run function tq:reshow"]
    fn["reshow"] = ["scoreboard players set @s tq_wait 0", "function tq:show_q"]

    # ---------------- sidebar
    fn["sidebar"] = [
        "scoreboard players operation #m tq_cfg = #interval tq_cfg",
        "scoreboard players operation #m tq_cfg -= @s tq_timer",
        "scoreboard players operation #m tq_cfg /= #1200 tq_cfg",
        "scoreboard players add #m tq_cfg 1",
        "execute unless score @s tq_state matches 0 run scoreboard players set #m tq_cfg 0",
    ]
    if t["numberformat"]:
        fn["sidebar"] += [
            "execute store result storage tq:tmp c int 1 run scoreboard players get @s tq_right",
            "execute store result storage tq:tmp w int 1 run scoreboard players get @s tq_wrong",
            "execute store result storage tq:tmp m int 1 run scoreboard players get #m tq_cfg",
            "function tq:sidebar_names with storage tq:tmp",
        ]
        fn["sidebar_names"] = [
            '$scoreboard players display name Nxt tq_score {"text":"Next question: $(m) min","color":"aqua"}',
            '$scoreboard players display name Cor tq_score {"text":"Correct: $(c)","color":"green"}',
            '$scoreboard players display name Wrg tq_score {"text":"Wrong: $(w)","color":"red"}',
        ]
    else:
        fn["sidebar"] += [
            "scoreboard players operation minutes tq_score = #m tq_cfg",
            "scoreboard players operation Correct tq_score = @s tq_right",
            "scoreboard players operation Wrong tq_score = @s tq_wrong",
        ]
    settings_functions(fn, t, qs, counts, rnames, pnames, triggers, reset_trigger)
    effect_functions(fn, t, rewards, punishments)
    return fn

# ----------------------------------------------------------------------------- functions: settings + UI
def settings_functions(fn, t, qs, counts, rnames, pnames, triggers, reset_trigger):
    iv = score("#interval_min", "tq_cfg")
    fn["set_diff"] = [   # #tmp = requested value
        "execute if score #tmp tq_cfg matches 2.. run scoreboard players set #diff tq_cfg 2",
        "execute if score #tmp tq_cfg matches ..1 run scoreboard players set #diff tq_cfg 1",
        "execute if score #diff tq_cfg matches 1 run " + bar("Difficulty: NORMAL", "green"),
        "execute if score #diff tq_cfg matches 2 run " + bar("Difficulty: HARD", "red"),
        "execute if score #diff tq_cfg matches 1 run " + msg("difficulty: NORMAL.", "green"),
        "execute if score #diff tq_cfg matches 2 run " + msg("difficulty: HARD.", "red"),
    ]
    fn["set_interval"] = [   # #interval_min = requested minutes
        "execute if score #interval_min tq_cfg matches ..0 run scoreboard players set #interval_min tq_cfg 1",
        f"execute if score #interval_min tq_cfg matches {INTERVAL_MAX_MIN + 1}.. run scoreboard players set #interval_min tq_cfg {INTERVAL_MAX_MIN}",
        "scoreboard players operation #interval tq_cfg = #interval_min tq_cfg",
        "scoreboard players operation #interval tq_cfg *= #1200 tq_cfg",
        "function tq:interval_msg",
    ]
    fn["interval_msg"] = [
        "title @s actionbar " + jcomp([comp("Interval: a question every ", "aqua"), dict(iv, color="aqua"), comp(" min", "aqua")]),
        tell([comp(TAG, "gold"), comp("a question every ", "gray"), dict(iv, color="gray"), comp(" minute(s).", "gray")]),
    ]
    fn["toggle_stats"] = [
        "scoreboard players operation #tmp tq_cfg = #sidebar tq_cfg",
        "execute if score #tmp tq_cfg matches 0 run scoreboard players set #sidebar tq_cfg 1",
        "execute if score #tmp tq_cfg matches 1 run scoreboard players set #sidebar tq_cfg 0",
        "execute if score #sidebar tq_cfg matches 1 run scoreboard objectives setdisplay sidebar tq_score",
        "execute if score #sidebar tq_cfg matches 0 run scoreboard objectives setdisplay sidebar",
        "execute if score #sidebar tq_cfg matches 1 run function tq:sidebar",
        "execute if score #sidebar tq_cfg matches 1 run " + bar("Score display: ON", "green"),
        "execute if score #sidebar tq_cfg matches 0 run " + bar("Score display: OFF", "red"),
        "execute if score #sidebar tq_cfg matches 1 run " + msg("score display ON.", "green"),
        "execute if score #sidebar tq_cfg matches 0 run " + msg("score display off.", "gray"),
    ]
    for prefix, names, kind in (("r", rnames, "Reward"), ("p", pnames, "Punishment")):
        for i, name in enumerate(names, 1):
            n = f"{prefix}{i}"
            fn[f"tog/{n}"] = [
                f"scoreboard players operation #f tq_cfg = {n} tq_on",
                f"execute if score #f tq_cfg matches 1 run scoreboard players set {n} tq_on 0",
                f"execute if score #f tq_cfg matches 0 run scoreboard players set {n} tq_on 1",
                f"execute if score {n} tq_on matches 1 run " + bar(f"{kind}: {name} is ON", "green"),
                f"execute if score {n} tq_on matches 0 run " + bar(f"{kind}: {name} is OFF", "red"),
                f"execute if score {n} tq_on matches 1 run " + msg(f"{kind} {name}: ON", "green"),
                f"execute if score {n} tq_on matches 0 run " + msg(f"{kind} {name}: OFF", "red"),
            ]
        fn[f"all_{prefix}_on"] = [f"scoreboard players set {prefix}{i} tq_on 1" for i in range(1, len(names) + 1)] + [
            bar(f"Every {kind.lower()} is ON", "green"), msg(f"every {kind.lower()} is ON.", "green")]
        fn[f"all_{prefix}_off"] = [f"scoreboard players set {prefix}{i} tq_on 0" for i in range(1, len(names) + 1)] + [
            bar(f"Every {kind.lower()} is OFF", "red"), msg(f"every {kind.lower()} is OFF - that side does nothing now.", "red")]

    # ---------------- chat triggers (usable by hand: /trigger tq_x ...)
    fn["trig/ask"] = reset_trigger("tq_ask") + [
        "execute if score @s tq_state matches 0 run function tq:ask",
        "execute unless score @s tq_state matches 0 run function tq:reshow",
    ]
    fn["trig/interval"] = ["scoreboard players operation #interval_min tq_cfg = @s tq_interval"] + reset_trigger("tq_interval") + ["function tq:set_interval"]
    fn["trig/diff"] = ["scoreboard players operation #tmp tq_cfg = @s tq_diff"] + reset_trigger("tq_diff") + ["function tq:set_diff"]
    fn["trig/stats"] = reset_trigger("tq_stats") + ["function tq:toggle_stats"]
    fn["trig/menu"] = reset_trigger("tq_menu") + ["function tq:ui/menu_show"]
    fn["trig/opts"] = reset_trigger("tq_opts") + ["function tq:ui/options_show"]
    fn["trig/rl"] = reset_trigger("tq_rl") + ["function tq:ui/rewards_show"]
    fn["trig/pl"] = reset_trigger("tq_pl") + ["function tq:ui/punishments_show"]
    for trig, prefix, names in (("tq_tr", "r", rnames), ("tq_tp", "p", pnames)):
        fn[f"trig/{trig[3:]}"] = [f"scoreboard players operation #tmp tq_cfg = @s {trig}"] + reset_trigger(trig) + \
            [f"execute if score #tmp tq_cfg matches {i} run function tq:tog/{prefix}{i}" for i in range(1, len(names) + 1)] + \
            [f"execute if score #tmp tq_cfg matches {ALL_ON} run function tq:all_{prefix}_on",
             f"execute if score #tmp tq_cfg matches {ALL_OFF} run function tq:all_{prefix}_off"]
    status = reset_trigger("tq_status") + [
        "scoreboard players set #off tq_cfg 0",
        msg("---- status ----", "gold"),
        "execute if score #diff tq_cfg matches 1 run " + msg("difficulty: NORMAL.", "gray"),
        "execute if score #diff tq_cfg matches 2 run " + msg("difficulty: HARD.", "gray"),
        "function tq:interval_msg",
        "execute if score #sidebar tq_cfg matches 1 run " + msg("score display: on.", "gray"),
        "execute if score #sidebar tq_cfg matches 0 run " + msg("score display: off.", "gray"),
    ]
    for prefix, names, kind in (("r", rnames, "reward"), ("p", pnames, "punishment")):
        for i, name in enumerate(names, 1):
            status.append(f"execute if score {prefix}{i} tq_on matches 0 run scoreboard players add #off tq_cfg 1")
            status.append(f"execute if score {prefix}{i} tq_on matches 0 run " + msg(f"{kind} OFF: {name}", "red"))
    status.append("execute if score #off tq_cfg matches 0 run " + msg("every reward and punishment is ON.", "green"))
    fn["trig/status"] = status

    # ---------------- the screens
    if not t["dialogs"]:
        fn["ui/menu_show"] = chat_menu(counts)
        fn["ui/options_show"] = chat_options()
        fn["ui/rewards_show"] = chat_list("Rewards", "r", rnames, "tq_tr", "tq_rl")
        fn["ui/punishments_show"] = chat_list("Punishments", "p", pnames, "tq_tp", "tq_pl")
        return
    # 26: inline dialogs built by macro functions; dialog buttons change something, then re-show the screen
    fn["trig/iv"] = ["scoreboard players operation #interval_min tq_cfg = @s tq_iv"] + reset_trigger("tq_iv") + [
        "function tq:set_interval", "function tq:ui/options_show"]
    fn["trig/ui"] = ["scoreboard players operation #tmp tq_cfg = @s tq_ui"] + reset_trigger("tq_ui") + [
        f"execute if score #tmp tq_cfg matches {UI_STATS} run function tq:ui/toggle_stats",
        f"execute if score #tmp tq_cfg matches {UI_STATS_OPT} run function tq:ui/toggle_stats_opt",
        f"execute if score #tmp tq_cfg matches {UI_DIFF} run function tq:ui/toggle_diff",
        f"execute if score #tmp tq_cfg matches {UI_CLOSE} run dialog clear @s",
        f"execute if score #tmp tq_cfg matches {UI_ALL_R} run function tq:ui/all_r_on",
        f"execute if score #tmp tq_cfg matches {UI_ALL_R_OFF} run function tq:ui/all_r_off",
        f"execute if score #tmp tq_cfg matches {UI_ALL_P} run function tq:ui/all_p_on",
        f"execute if score #tmp tq_cfg matches {UI_ALL_P_OFF} run function tq:ui/all_p_off",
        f"execute if score #tmp tq_cfg matches {UI_R_BASE + 1}..{UI_R_BASE + len(rnames)} run function tq:ui/tog_r",
        f"execute if score #tmp tq_cfg matches {UI_P_BASE + 1}..{UI_P_BASE + len(pnames)} run function tq:ui/tog_p",
    ]
    fn["ui/toggle_diff"] = ["scoreboard players set #tmp tq_cfg 1",
                            "execute if score #diff tq_cfg matches 1 run scoreboard players set #tmp tq_cfg 2",
                            "function tq:set_diff", "function tq:ui/options_show"]
    fn["ui/toggle_stats"] = ["function tq:toggle_stats", "function tq:ui/menu_show"]
    fn["ui/toggle_stats_opt"] = ["function tq:toggle_stats", "function tq:ui/options_show"]
    for prefix, names, kind, base in (("r", rnames, "rewards", UI_R_BASE), ("p", pnames, "punishments", UI_P_BASE)):
        fn[f"ui/all_{prefix}_on"] = [f"function tq:all_{prefix}_on", f"function tq:ui/{kind}_show"]
        fn[f"ui/all_{prefix}_off"] = [f"function tq:all_{prefix}_off", f"function tq:ui/{kind}_show"]
        fn[f"ui/tog_{prefix}"] = [f"scoreboard players remove #tmp tq_cfg {base}"] + \
            [f"execute if score #tmp tq_cfg matches {i} run function tq:tog/{prefix}{i}" for i in range(1, len(names) + 1)] + \
            [f"function tq:ui/{kind}_show"]
        show = []
        for i in range(1, len(names) + 1):
            show.append(f'execute if score {prefix}{i} tq_on matches 1 run data modify storage tq:tmp c{i} set value "green"')
            show.append(f'execute unless score {prefix}{i} tq_on matches 1 run data modify storage tq:tmp c{i} set value "red"')
        fn[f"ui/{kind}_show"] = show + [f"function tq:ui/{kind} with storage tq:tmp"]
    fn["ui/rewards"] = [ui_macro(toggle_dialog("Rewards", rnames, UI_R_BASE, UI_ALL_R, UI_ALL_R_OFF))]
    fn["ui/punishments"] = [ui_macro(toggle_dialog("Punishments", pnames, UI_P_BASE, UI_ALL_P, UI_ALL_P_OFF))]
    fn["ui/state"] = [   # the live values the screens show: sd/sdc score display, dn/dc difficulty, iv interval
        'execute if score #sidebar tq_cfg matches 1 run data modify storage tq:tmp sd set value "ON"',
        'execute if score #sidebar tq_cfg matches 1 run data modify storage tq:tmp sdc set value "green"',
        'execute unless score #sidebar tq_cfg matches 1 run data modify storage tq:tmp sd set value "OFF"',
        'execute unless score #sidebar tq_cfg matches 1 run data modify storage tq:tmp sdc set value "red"',
        'execute if score #diff tq_cfg matches 2 run data modify storage tq:tmp dn set value "HARD"',
        'execute if score #diff tq_cfg matches 2 run data modify storage tq:tmp dc set value "red"',
        'execute unless score #diff tq_cfg matches 2 run data modify storage tq:tmp dn set value "NORMAL"',
        'execute unless score #diff tq_cfg matches 2 run data modify storage tq:tmp dc set value "green"',
        "execute store result storage tq:tmp iv int 1 run scoreboard players get #interval_min tq_cfg",
        'data modify storage tq:tmp tpl set value "trigger tq_iv set $(iv)"',
    ]
    fn["ui/menu_show"] = ["function tq:ui/state", "function tq:ui/menu with storage tq:tmp"]
    fn["ui/menu"] = [ui_macro(menu_dialog(counts))]
    fn["ui/options_show"] = ["function tq:ui/state", "function tq:ui/options with storage tq:tmp"]
    fn["ui/options"] = [ui_macro(options_dialog())]

# ----------------------------------------------------------------------------- functions: rewards, punishments
MOB_OFFSETS = [(4, 0), (-4, 0), (0, 4), (0, -4), (3, 3), (-3, 3), (3, -3), (-3, -3), (6, 2), (-6, -2), (2, 6), (-2, -6)]
MOBS = ["minecraft:zombie", "minecraft:skeleton", "minecraft:creeper", "minecraft:spider", "minecraft:wither_skeleton",
        "minecraft:creeper ~ ~ ~ {powered:1b}"]   # 6 = charged creeper (Creeper ambush)
SCATTER = [(0.35, 0.0), (0.25, 0.25), (0.0, 0.35), (-0.25, 0.25), (-0.35, 0.0), (-0.25, -0.25), (0.0, -0.35), (0.25, -0.25), (0.2, 0.3)]


def weighted_roll(fn, prefix, keys, out, kind):
    """Cumulative-weight roll over fake players <prefix><i> in tq_w (only the ones with tq_on = 1); the chosen
    weight is halved so repeats get rarer."""
    names = [f"{prefix}{i}" for i in range(1, len(keys) + 1)]
    fn[f"{out}/random"] = ["scoreboard players set #sum tq_cfg 0"] + [
        f"execute if score {n} tq_on matches 1 run scoreboard players operation #sum tq_cfg += {n} tq_w" for n in names] + [
        "execute if score #sum tq_cfg matches ..0 run " + msg(f"every {kind} is switched off - nothing happens.", "yellow"),
        f"execute if score #sum tq_cfg matches 1.. run function tq:{out}/roll",
    ]
    lines = [
        "function tq:rng",
        "scoreboard players operation #roll tq_cfg = #rnd tq_cfg",
        "scoreboard players operation #roll tq_cfg %= #sum tq_cfg",
        "scoreboard players add #roll tq_cfg 1",
        "scoreboard players set #acc tq_cfg 0",
        "scoreboard players set #pick tq_cfg 0",
    ]
    for i, n in enumerate(names, 1):
        lines.append(f"execute if score #pick tq_cfg matches 0 if score {n} tq_on matches 1 run scoreboard players operation #acc tq_cfg += {n} tq_w")
        lines.append(f"execute if score #pick tq_cfg matches 0 if score {n} tq_on matches 1 if score #roll tq_cfg <= #acc tq_cfg run scoreboard players set #pick tq_cfg {i}")
    for i, n in enumerate(names, 1):
        lines.append(f"execute if score #pick tq_cfg matches {i} run scoreboard players operation {n} tq_w /= #two tq_cfg")
        lines.append(f"execute if score #pick tq_cfg matches {i} if score {n} tq_w matches ..0 run scoreboard players set {n} tq_w 1")
    for i, key in enumerate(keys, 1):
        lines.append(f"execute if score #pick tq_cfg matches {i} run function tq:{out}/{key}")
    fn[f"{out}/roll"] = lines


def effect_functions(fn, t, rewards, punishments):
    A = t["attr"]
    rkeys, pkeys = [r[0] for r in rewards], [p[0] for p in punishments]
    weighted_roll(fn, "r", rkeys, "reward", "reward")
    weighted_roll(fn, "p", pkeys, "penalty", "punishment")
    for key, short, title, subtitle, color, _, cmds in rewards:
        fn[f"reward/{key}"] = ["title @s times 10 70 20", sub(subtitle, "white")] + cmds + [
            f"title @s title {jcomp(comp(title, color))}",
            tell([comp(TAG, "gold"), comp(f"Reward: {short} - {subtitle}", "green")])]
    fn["reward/heart_set"] = [f"execute if score #h tq_cfg matches {h} run attribute @s {A['max_health']} base set {h}" for h in range(22, 41, 2)]

    # freeze: every mob nearby loses its AI for FREEZE_SECONDS; countdown on the action bar
    fn["freeze/start"] = [
        f"scoreboard players set #freeze tq_cfg {FREEZE_SECONDS}",
        "execute as @e[type=!player,type=!item,type=!experience_orb,distance=..96] run function tq:freeze/hold",
        "function tq:freeze/bar",
    ]
    fn["freeze/hold"] = ["execute unless data entity @s {NoAI:1b} run tag @s add tq_frozen",
                         "execute if entity @s[tag=tq_frozen] run data merge entity @s {NoAI:1b}"]
    fn["freeze/release"] = ["data merge entity @s {NoAI:0b}", "tag @s remove tq_frozen"]
    fn["freeze/tick"] = [
        "scoreboard players remove #freeze tq_cfg 1",
        "execute if score #freeze tq_cfg matches 1.. run function tq:freeze/bar",
        "execute if score #freeze tq_cfg matches ..0 run function tq:freeze/stop",
    ]
    fn["freeze/bar"] = ["title @a actionbar " + jcomp([comp("TIME FREEZE  ", "aqua", bold=True), score("#freeze", "tq_cfg", "aqua")])]
    fn["freeze/stop"] = ["scoreboard players set #freeze tq_cfg 0",
                         "execute as @e[tag=tq_frozen] run function tq:freeze/release", bar("time resumes", "gray", "@a")]
    fn["freeze/stop_silent"] = ["scoreboard players set #freeze tq_cfg 0", "execute as @e[tag=tq_frozen] run function tq:freeze/release"]

    for key, short, title, detail, _, cmds in punishments:
        fn[f"penalty/{key}"] = ["title @s times 10 70 20"] + cmds + [
            f"title @s title {jcomp(comp(title, 'dark_red'))}",
            bar(detail, "red"),
            tell([comp(TAG, "gold"), comp(f"Punishment: {short} - {detail}.  ", "red"), comp("Correct: ", "yellow"),
                  {"nbt": "ansfull", "storage": "tq:tmp", "color": "white"}])]
    fn["penalty/starve2"] = ["effect give @s minecraft:hunger 180 2 true"]
    lost = ["function tq:rng", "scoreboard players operation #r tq_cfg = #rnd tq_cfg", "scoreboard players operation #r tq_cfg %= #eight tq_cfg"]
    for i, (dx, dz) in enumerate([(0, -300), (0, 300), (-300, 0), (300, 0), (212, 212), (-212, 212), (212, -212), (-212, -212)]):
        lost.append(f"execute if score #r tq_cfg matches {i} if predicate tq:overworld run spreadplayers ~{dx} ~{dz} 0 12 false @s")
        lost.append(f"execute if score #r tq_cfg matches {i} unless predicate tq:overworld run spreadplayers ~{dx} ~{dz} 0 12 under 120 false @s")
    fn["penalty/lost_go"] = lost

    # mobs: pick one of 12 spots around the player, the first free height wins, otherwise right at the player
    fn["penalty/mob_loop"] = ["scoreboard players add #i tq_cfg 1", "function tq:penalty/mob_one",
                              "execute if score #i tq_cfg matches ..19 run function tq:penalty/mob_loop"]
    fn["penalty/mob_one"] = ["function tq:rng", "scoreboard players operation #mob tq_cfg = #rnd tq_cfg",
                             "scoreboard players operation #mob tq_cfg %= #five tq_cfg", "scoreboard players add #mob tq_cfg 1",
                             "function tq:penalty/mob_place"]
    fn["penalty/creeper_loop"] = ["scoreboard players add #i tq_cfg 1", "scoreboard players set #mob tq_cfg 6",
                                  "function tq:penalty/mob_place",
                                  "execute if score #i tq_cfg matches ..3 run function tq:penalty/creeper_loop"]
    fn["penalty/mob_place"] = [
        "function tq:rng", "scoreboard players operation #r tq_cfg = #rnd tq_cfg", "scoreboard players operation #r tq_cfg %= #twelve tq_cfg",
        "scoreboard players set #placed tq_cfg 0",
    ] + [f"execute if score #r tq_cfg matches {i} positioned ~{dx} ~ ~{dz} run function tq:penalty/mob_at" for i, (dx, dz) in enumerate(MOB_OFFSETS)] + [
        "execute if score #placed tq_cfg matches 0 run function tq:penalty/mob_summon"]
    fn["penalty/mob_at"] = [f"execute if score #placed tq_cfg matches 0 positioned ~ ~{dy} ~ run function tq:penalty/mob_try" for dy in (0, 1, -1, 2, -2)]
    fn["penalty/mob_try"] = ["scoreboard players set #free tq_cfg 0"] + [
        f"execute if block ~ ~ ~ minecraft:{a} if block ~ ~1 ~ minecraft:{b} run scoreboard players set #free tq_cfg 1"
        for a in ("air", "cave_air") for b in ("air", "cave_air")] + [
        "execute if score #free tq_cfg matches 1 run function tq:penalty/mob_summon"]
    fn["penalty/mob_summon"] = [
        f"execute if score #mob tq_cfg matches {i} store success score #placed tq_cfg run summon {m}" + ("" if "~" in m else " ~ ~ ~")
        for i, m in enumerate(MOBS, 1)]
    fn["penalty/phantoms"] = [f"summon minecraft:phantom ~{dx} ~18 ~{dz}" for dx, dz in ((0, 0), (4, 3), (-4, 3), (3, -4), (-3, -4), (0, 6))]

    def drop_cmds(path, slot, mx, mz):
        it = "@e[type=item,tag=tq_drop,limit=1,sort=nearest]"
        ck, ek, cv = t["count_key"], t["extra_key"], t["count_val"]
        return [
            f"execute if data entity @s {path} run summon minecraft:item ~ ~1 ~ {{Tags:[\"tq_drop\"],PickupDelay:60s,Motion:[{mx}d,0.35d,{mz}d],Item:{{id:\"minecraft:stone\",{ck}:{cv}}}}}",
            f"execute if data entity @s {path} run data modify entity {it} Item.id set from entity @s {path}.id",
            f"execute if data entity @s {path} run data modify entity {it} Item.{ck} set from entity @s {path}.{ck}",
            f"execute if data entity @s {path}.{ek} run data modify entity {it} Item.{ek} set from entity @s {path}.{ek}",
            f"execute if data entity @s {path} run " + t["item_replace"](slot, "minecraft:air"),
            "tag @e[type=item,tag=tq_drop] remove tq_drop",
        ]
    drop = ["playsound minecraft:entity.item.pickup master @s ~ ~ ~ 1 0.5"]
    for i in range(9):
        drop += drop_cmds(f"Inventory[{{Slot:{i}b}}]", f"hotbar.{i}", *SCATTER[i])
    fn["penalty/drop_hotbar"] = drop
    drop = ["playsound minecraft:item.armor.equip_generic master @s ~ ~ ~ 1 0.5"]
    for k, part in enumerate(("head", "chest", "legs", "feet")):
        drop += drop_cmds(t["armor_path"](part), f"armor.{part}", *SCATTER[k * 2])
    fn["penalty/drop_armor"] = drop
    if t["macros"]:
        fn["penalty/to_spawn"] = ["execute unless data storage tq:spawn x run tp @s ~ ~200 ~",
                                  "execute if data storage tq:spawn x run function tq:penalty/to_spawn_tp with storage tq:spawn"]
        fn["penalty/to_spawn_tp"] = ["$execute in minecraft:overworld run tp @s $(x) $(y) $(z)"]
    else:
        aec = "@e[type=minecraft:area_effect_cloud,tag=tq_spawn"
        fn["penalty/to_spawn"] = [f"execute in minecraft:overworld if entity {aec}] run tp @s {aec},limit=1]",
                                  f"execute in minecraft:overworld unless entity {aec}] run tp @s ~ ~200 ~"]
    fn["penalty/buried_go"] = [
        "execute store result score #y tq_cfg run data get entity @s Pos[1]",
        "scoreboard players set #ok tq_cfg 0",
        "execute if predicate tq:overworld if score #y tq_cfg matches -40.. run scoreboard players set #ok tq_cfg 1",
        "execute unless predicate tq:overworld if score #y tq_cfg matches 25.. run scoreboard players set #ok tq_cfg 1",
        "execute if score #ok tq_cfg matches 1 run tp @s ~ ~-15 ~",
        "execute if score #ok tq_cfg matches 0 run function tq:penalty/lost_go",
    ]

    # ---------------- admin helpers (need cheats)
    fn["admin/test_reward"] = ["function tq:reward/random"]
    fn["admin/test_penalty"] = ["data remove storage tq:tmp ans", "data remove storage tq:tmp ansfull",
                                'title @s subtitle ""', "function tq:penalty/random"]
    fn["admin/reset_seen"] = ["function tq:reset_n", "function tq:reset_h"]
    fn["admin/reset_weights"] = [f"scoreboard players set r{i} tq_w 64" for i in range(1, len(rkeys) + 1)] + \
        [f"scoreboard players set p{i} tq_w 64" for i in range(1, len(pkeys) + 1)] + [msg("weights reset.", "gold")]
    fn["admin/enable_all"] = ["function tq:all_r_on", "function tq:all_p_on"]
    G = t["gamerule"]
    fn["admin/feedback_on"] = [f"gamerule {G['feedback']} true", f"gamerule {G['log']} true", "scoreboard players set #fb_set tq_cfg 2",
                               msg("command feedback is back on (the Triggered [...] lines will show again).", "gold")]
    fn["admin/clear"] = ["effect clear @s", "function tq:timers/clear_all", "function tq:freeze/stop_silent", msg("cleared.", "gold")]

# ----------------------------------------------------------------------------- packaging
def write(path, content):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if isinstance(content, (dict, list)):
        content = json.dumps(content, ensure_ascii=True, separators=(",", ":"))
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)


def zip_dir(src, dest, top_only=None):
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(src):
            for f in sorted(files):
                p = os.path.join(root, f)
                rel = os.path.relpath(p, src).replace("\\", "/")
                if top_only is None or rel.split("/")[0] in top_only:
                    z.write(p, rel)


def build_target(name, qs, ranges, known):
    t = TARGETS[name]
    stage = os.path.join(BUILD, f"stage_{name}")
    if os.path.exists(stage):
        shutil.rmtree(stage)
    data = os.path.join(stage, "data", NS)
    if t["dialogs"]:
        for q in qs:
            write(os.path.join(data, "dialog", f"q_{q['n']}.json"), question_dialog(q))
    fns = functions(t, qs, ranges)
    for fname, lines in fns.items():
        write(os.path.join(data, t["fn_dir"], fname + ".mcfunction"), "\n".join(lines) + "\n")
    write(os.path.join(stage, "data", "minecraft", t["tag_dir"], "load.json"), {"values": [f"{NS}:load"]})
    write(os.path.join(stage, "data", "minecraft", t["tag_dir"], "tick.json"), {"values": [f"{NS}:tick"]})
    lt, unknown = useful_loot_table(t, known)
    write(os.path.join(data, t["loot_dir"], "useful_item.json"), lt)
    write(os.path.join(data, t["pred_dir"], "overworld.json"),
          {"condition": "minecraft:location_check", "predicate": {"dimension": "minecraft:overworld"}})
    write(os.path.join(stage, "pack.mcmeta"), {"pack": dict(
        {"description": f"{NAME} {VERSION} - {len(qs)} Minecraft trivia questions, rewards and punishments ({t['label']})",
         "pack_format": t["pack_format"]}, **t["mcmeta_extra"])})
    os.makedirs(DIST, exist_ok=True)
    zip_dir(stage, os.path.join(DIST, f"TriviaChallenge-{name}.zip"), top_only=("pack.mcmeta", "data"))
    if t["jar"]:
        write(os.path.join(stage, "META-INF", "mods.toml"),
              'modLoader="lowcodefml"\nloaderVersion="[40,)"\nlicense="MIT"\nshowAsResourcePack=false\n[[mods]]\n'
              f'modId="{MODID}"\nversion="{VERSION}"\ndisplayName="{NAME}"\nauthors="W1ndLeaf"\n'
              f"description='''{NAME}: Minecraft trivia with rewards and punishments.'''\n")
        write(os.path.join(stage, "META-INF", "neoforge.mods.toml"),
              'modLoader="lowcodefml"\nloaderVersion="[1,)"\nlicense="MIT"\n[[mods]]\n'
              f'modId="{MODID}"\nversion="{VERSION}"\ndisplayName="{NAME}"\ndescription=\'\'\'{NAME}\'\'\'\n')
        write(os.path.join(stage, "fabric.mod.json"),
              {"schemaVersion": 1, "id": MODID, "version": VERSION, "name": NAME, "environment": "*",
               "depends": {"fabric-resource-loader-v0": "*"}})
        zip_dir(stage, os.path.join(DIST, f"TriviaChallenge-{name}.jar"))
    problems = lint(stage, t)
    return {"functions": len(fns), "loot": len(lt["pools"][0]["entries"]), "unknown": unknown,
            "rewards": len(REWARDS(t)), "punishments": len(PUNISHMENTS(t)), "problems": problems}


# ----------------------------------------------------------------------------- lint
KNOWN = {"scoreboard", "execute", "function", "tellraw", "title", "dialog", "tag", "effect", "attribute", "give", "loot",
         "summon", "tp", "spawnpoint", "time", "playsound", "item", "data", "kill", "spreadplayers", "random", "gamerule",
         "team", "replaceitem"}
JSON_CMDS = re.compile(r"^(tellraw \S+ |title \S+ (?:title|subtitle|actionbar) |team modify \S+ prefix |"
                       r"scoreboard objectives add \S+ dummy |scoreboard players display name \S+ \S+ )")
DUMMY = {"iv": "5", "tpl": "trigger tq_iv set 5"}


def lint(stage, t):
    problems = []
    fdir = os.path.join(stage, "data", NS, t["fn_dir"])
    for root, _, files in os.walk(fdir):
        for f in files:
            rel = os.path.relpath(os.path.join(root, f), fdir).replace("\\", "/")
            for ln, line in enumerate(open(os.path.join(root, f), encoding="utf-8"), 1):
                line = line.rstrip("\n")
                if not line.strip():
                    continue
                macro = line.startswith("$")
                body = line[1:] if macro else line
                if macro and not t["macros"]:
                    problems.append(f"{rel}:{ln} macro line in a target without macros")
                if macro and "$(" not in body:
                    problems.append(f"{rel}:{ln} macro line without $(...)")
                if not macro and "$(" in body and 'set value "trigger tq_iv set $(iv)"' not in body:
                    problems.append(f"{rel}:{ln} $(...) in a non-macro line")
                head = body.split()[0]
                if head not in KNOWN:
                    problems.append(f"{rel}:{ln} unknown command '{head}'")
                for a, b in (("{", "}"), ("[", "]"), ("(", ")")):
                    if body.count(a) != body.count(b):
                        problems.append(f"{rel}:{ln} unbalanced {a}{b}: {body[:80]}")
                if body.count('"') % 2:
                    problems.append(f"{rel}:{ln} odd number of double quotes: {body[:80]}")
                if any(ord(c) > 127 for c in body):
                    problems.append(f"{rel}:{ln} non-ascii")
                for bad in t["forbidden"]:
                    if bad in body and not (bad == "$(" and macro):
                        problems.append(f"{rel}:{ln} '{bad}' does not exist in {t['label']}")
                if rel == "load.mcfunction" and head not in ("scoreboard", "execute", "function", "tellraw", "team"):
                    problems.append(f"{rel}:{ln} keep load minimal - '{head}' belongs in a sub-function")
                for m in re.finditer(r"function tq:([a-z0-9_/]+)", body):
                    if not os.path.exists(os.path.join(fdir, m.group(1) + ".mcfunction")):
                        problems.append(f"{rel}:{ln} calls missing function tq:{m.group(1)}")
                for m in re.finditer(r"predicate tq:([a-z0-9_/]+)", body):
                    if not os.path.exists(os.path.join(stage, "data", NS, t["pred_dir"], m.group(1) + ".json")):
                        problems.append(f"{rel}:{ln} missing predicate {m.group(1)}")
                for m in re.finditer(r"dialog show @s tq:([a-z0-9_]+)", body):
                    if not os.path.exists(os.path.join(stage, "data", NS, "dialog", m.group(1) + ".json")):
                        problems.append(f"{rel}:{ln} shows missing dialog {m.group(1)}")
                text = re.sub(r"\$\((\w+)\)", lambda m: DUMMY.get(m.group(1), "green"), body)
                tail = body[len("execute "):] if body.startswith("execute ") else body
                tail = tail[tail.index(" run ") + 5:] if " run " in tail and body.startswith("execute ") else tail
                tail = re.sub(r"\$\((\w+)\)", lambda m: DUMMY.get(m.group(1), "green"), tail)
                jm = JSON_CMDS.match(tail)
                if jm:
                    try:
                        json.loads(tail[jm.end():])
                    except Exception as e:  # noqa
                        problems.append(f"{rel}:{ln} bad JSON text component: {e}")
                if macro and body.startswith("dialog show @s {"):
                    try:
                        d = json.loads(text[len("dialog show @s "):])
                    except Exception as e:  # noqa
                        problems.append(f"{rel}:{ln} inline dialog is not valid JSON: {e}")
                        continue
                    ex = d.get("exit_action", {}).get("action", {}).get("command")
                    if d.get("pause") is not False or d.get("after_action") != "none" or ex != f"trigger tq_ui set {UI_CLOSE}":
                        problems.append(f"{rel}:{ln} settings screen must be pause:false, after_action:none, footer = close trigger")
                    for a in d.get("actions", []):
                        cmd = a["action"].get("command", a["action"].get("template", ""))
                        if not cmd.startswith("trigger tq_"):
                            problems.append(f"{rel}:{ln} dialog button runs non-trigger command {cmd!r}")
    if t["dialogs"]:
        ddir = os.path.join(stage, "data", NS, "dialog")
        for f in os.listdir(ddir):
            d = json.load(open(os.path.join(ddir, f), encoding="utf-8"))
            for a in d.get("actions", []) + [d["exit_action"]]:
                if not a["action"]["command"].startswith("trigger tq_"):
                    problems.append(f"dialog {f}: button runs non-trigger command")
    json.load(open(os.path.join(stage, "data", NS, t["loot_dir"], "useful_item.json"), encoding="utf-8"))
    return problems


def apply_settings_sheet():
    """The Settings sheet of questions.xlsx overrides the DEFAULT_* constants: what every new world starts with.
    A missing sheet is created with the current defaults so it can be edited next time."""
    global DEFAULT_DIFF, DEFAULT_INTERVAL_MIN, DEFAULT_SIDEBAR, AUTO_MENU, DISABLED_REWARDS, DISABLED_PUNISHMENTS
    from questions_xlsx import read_settings, write_settings_sheet
    full = TARGETS["26"]   # the target that has every reward / punishment
    rewards = [(r[0], r[1]) for r in REWARDS(full)]
    punishments = [(p[0], p[1]) for p in PUNISHMENTS(full)]
    s = read_settings(XLSX, rewards, punishments)
    if s is None:
        write_settings_sheet(XLSX, rewards, punishments)
        print("added a Settings sheet to questions.xlsx (defaults)")
        return
    DEFAULT_DIFF, DEFAULT_INTERVAL_MIN, DEFAULT_SIDEBAR, AUTO_MENU = s["diff"], s["interval"], s["sidebar"], s["auto_menu"]
    DISABLED_REWARDS, DISABLED_PUNISHMENTS = s["off_rewards"], s["off_punishments"]
    print(f"settings: {'Hard' if DEFAULT_DIFF == 2 else 'Normal'}, every {DEFAULT_INTERVAL_MIN} min, score display "
          f"{'on' if DEFAULT_SIDEBAR else 'off'}, menu on first join {'on' if AUTO_MENU else 'off'}, "
          f"off: {DISABLED_REWARDS + DISABLED_PUNISHMENTS or 'nothing'}")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    targets = args or list(TARGETS)
    for x in targets:
        if x not in TARGETS:
            sys.exit(f"unknown target {x!r}; choose from {list(TARGETS)}")
    qs, ranges = load_questions()
    print(f"questions: {len(qs)}  (Normal {ranges['N']}, Hard {ranges['H']})  from {os.path.basename(XLSX)}")
    apply_settings_sheet()
    known = set(json.load(open(os.path.join(HERE, "nonblock_items.json")))) | BLOCK_ITEMS
    os.makedirs(BUILD, exist_ok=True)
    with open(os.path.join(BUILD, "questions_report.txt"), "w", encoding="utf-8") as f:
        for q in qs:
            f.write(f"{q['n']:4d}  {q['id']:8s} {q['pool']} {q['topic']:10s} {q['q'][:90]}\n")
    failed = False
    for name in targets:
        s = build_target(name, qs, ranges, known)
        z = os.path.join(DIST, f"TriviaChallenge-{name}.zip")
        print(f"[{name:4s}] {TARGETS[name]['label']}: {s['functions']} functions, {s['rewards']} rewards, "
              f"{s['punishments']} punishments, {s['loot']} items -> {os.path.relpath(z, ROOT)} ({os.path.getsize(z) // 1024} KB)"
              + (f"  !! unknown item ids: {s['unknown']}" if s["unknown"] else ""))
        for p in s["problems"][:40]:
            print("   -", p)
        failed |= bool(s["problems"])
    if "--copy" in sys.argv and os.path.isdir(INSTANCE_MODS) and "26" in targets:
        for old in glob.glob(os.path.join(INSTANCE_MODS, "TriviaChallenge-*.jar")):
            os.remove(old)
        shutil.copy2(os.path.join(DIST, "TriviaChallenge-26.jar"), INSTANCE_MODS)
        print("copied the 26 jar to", INSTANCE_MODS)
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
