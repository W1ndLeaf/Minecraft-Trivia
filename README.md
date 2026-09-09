# Trivia Challenge

Minecraft trivia that pops up while you play. Answer right and you get a random reward; answer wrong and you get a
random punishment. Made for speedrun / challenge streams. **Plain vanilla data pack - no mods, no resource pack.**

152 questions (76 Normal, 76 Hard), 18 rewards, 23 punishments, all editable in one Excel file.

## Download

Pick the zip for your Minecraft version from [`dist/`](dist/):

| Minecraft | File | Interface |
|---|---|---|
| **26.x** (also 1.21.11+) | [`dist/TriviaChallenge-26.zip`](dist/TriviaChallenge-26.zip) | pop-up dialog windows with buttons |
| **1.21 – 1.21.1** | [`dist/TriviaChallenge-1.21.zip`](dist/TriviaChallenge-1.21.zip) | clickable chat (dialogs only exist since 1.21.6) |
| **1.16.5** | [`dist/TriviaChallenge-1.16.zip`](dist/TriviaChallenge-1.16.zip) | clickable chat; a few effects that need newer game features are left out (see below) |

One code base builds all three, so the questions and the rules are identical. Do not rename the zips.

## Install

**New world:** Singleplayer → Create New World → *More* tab → **Data Packs** → **Open Pack Folder** → drop the zip into
the folder that opens → back in the game move *Trivia Challenge* from *Available* to *Selected* → Done → Create.

**Existing world:** put the zip into `.minecraft/saves/<your world>/datapacks/` and run `/reload` in the world
(or just re-enter the world). On Windows `.minecraft` is `%APPDATA%\.minecraft`.

**Server:** `<server folder>/world/datapacks/` and `/reload`.

That is all. When it is loaded you see `[Trivia] loaded 152 questions.` in chat.

Optional for modded players (Forge / NeoForge / Fabric+Fabric API): `dist/TriviaChallenge-26.jar` or
`dist/TriviaChallenge-1.21.jar` in `mods/` loads the same pack in every world automatically.

## How it plays

- On joining you get one chat line with the commands. `/trigger tq_menu` opens the menu (no cheats / OP needed).
- Every few minutes (default 5) a question appears with answers **A / B / C (/ D)**. Skipping counts as wrong.
- **Right = a random reward, wrong = a random punishment.** On screen you see what you got; when you were wrong the
  line under it says `Correct: B - <the right answer>`. A `[Trivia]` chat line keeps the record.
- Every reward and punishment starts with an **equal chance**; whatever hits you has its chance **halved**
  (64 → 32 → 16 …), so repeats get rarer and everything comes around. Every question is asked **exactly once**; when a
  pool is exhausted the pack says so and starts it over.
- **Score display** (sidebar): next question in N min, correct, wrong.
- **Options:** Normal / Hard, interval (1–30 min), score display on/off, and a list of every reward and punishment
  where each one can be switched off. Settings are saved in the world; a fresh world starts with the defaults.

### Commands (no cheats needed)

```
/trigger tq_menu               menu                 /trigger tq_opts        options
/trigger tq_ask                ask a question now   /trigger tq_status      print all settings
/trigger tq_diff set 1|2       Normal | Hard        /trigger tq_interval set 7   every 7 minutes
/trigger tq_stats              score display on/off
/trigger tq_rl / tq_pl         rewards / punishments list (switch items on/off)
```
With cheats: `/function tq:admin/test_reward`, `/function tq:admin/test_penalty`, `/function tq:admin/reset_seen`,
`/function tq:admin/reset_weights`, `/function tq:admin/enable_all`, `/function tq:admin/clear`,
`/function tq:admin/feedback_on` (the pack turns command feedback off so `/trigger` does not spam chat).

### Rewards

Classic tools (water bucket + boat) · Night Vision (permanent) · Slow Falling (permanent) · Safe spawn · +1 heart ·
Full belly · Second wind (Regen II 15 s) · 64 cobblestone · Shield · Iron pickaxe · **Time freeze** (every mob near you
stops for 10 s) · **Useful item** (one of 90) · Bow + 8 arrows · Fire Resistance 3 min · Haste 3 min · Speed 3 min ·
Long arms (reach +2, 3 min)* · Tiny (half size, 3 min)*

### Punishments

Grounded (no jumping 2 min)* · Giant (2×, 3 min)* · Short arms (reach 2, 3 min)* · Glass bones (fall damage ×3, 5 min)* ·
Slowness II 2 min · Mining Fatigue 3 min · Levitation V 8 s · Lost (300 blocks away) · Mob wave (20 hostile mobs) ·
A Warden* · Sky drop (200 up) · The Wither · Starvation · Lights out (Blindness 1 min) · Butterfingers (hotbar dropped) ·
Back to spawn · Poison II 30 s · Feeble (Weakness II 3 min) · Phantom night · Creeper ambush (4 charged creepers) ·
Armor strip · Tool break · Buried alive (15 blocks down)

\* needs game features that 1.16 does not have - these are simply not in the 1.16 pack. Not a bug: you cannot sprint
while blind or with 3 hunger shanks or less (vanilla rules).

## Editing the questions

Everything comes from **`questions.xlsx`** (sheet *Questions*): one row per question - ID, Pool (Normal/Hard),
Topic, Question, answers A–D (C and D may be empty), Correct (A–D), Explanation, Source, Version. Add rows, delete
rows, change text; the *How to edit* sheet inside the file explains the rules. Then rebuild:

```
pip install openpyxl        (once)
python tools/build.py       (all versions)      python tools/build.py 26     (one version)
```

The new zips land in `dist/`. The build checks every row and stops with the row number if something is wrong
(missing answer, bad letter, duplicate ID).

## Settings for every world

A data pack can only save things inside one world - a new world always starts from scratch, and nothing in vanilla
can carry settings from world to world. So the settings that should apply **everywhere** are baked into the zip:
open the **Settings** sheet of `questions.xlsx`, set difficulty, interval, score display, menu-on-first-join and
On/Off for every single reward and punishment, save, rebuild. Every world you then create, on every version, starts
exactly like that. Changes made in game (menu / options) still apply to that world only.

## Repository layout

```
questions.xlsx      the question bank (edit this)
dist/               ready-to-use packs, one per Minecraft version
tools/build.py      generator + checker (Python 3, openpyxl)
tools/questions_xlsx.py
```
