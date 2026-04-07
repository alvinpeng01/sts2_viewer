#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
import os
import sys
import glob
import json
from collections import defaultdict

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_FILE = os.path.join(BASE_DIR, "sts2_cards.xlsx")

CLASS_COLORS = {
    "IRONCLAD": "FF8B8B",
    "SILENT": "8BFF8B",
    "DEFECT": "8B8BFF",
    "NECROBINDER": "FF8BFF",
    "REGENT": "FFFF8B",
    "COLORLESS": "E0E0E0",
}

CARD_CLASSIFICATIONS = {
    "ABRASIVE": "SILENT",
    "ACCELERANT": "SILENT",
    "ACCURACY": "SILENT",
    "ACROBATICS": "SILENT",
    "ADAPTIVE_STRIKE": "DEFECT",
    "ADRENALINE": "SILENT",
    "AFTERIMAGE": "SILENT",
    "AGGRESSION": "IRONCLAD",
    "ALCHEMIZE": "COLORLESS",
    "ALL_FOR_ONE": "DEFECT",
    "ANGER": "IRONCLAD",
    "ANOINTED": "DEFECT",
    "ANTICIPATE": "DEFECT",
    "APOTHEOSIS": "COLORLESS",
    "ARBITER": "NECROBINDER",
    "ARMAMENTS": "IRONCLAD",
    "ARTIFACT": "DEFECT",
    "ASH": "REGENT",
    "BACKFLIP": "SILENT",
    "BALANCE": "REGENT",
    "BANDAGE_UP": "COLORLESS",
    "BANE": "SILENT",
    "BASH": "IRONCLAD",
    "BEAM_CELL": "DEFECT",
    "BEAT_DOWN": "IRONCLAD",
    "BECKON": "NECROBINDER",
    "BERSERK": "IRONCLAD",
    "BESTIAL": "REGENT",
    "BET": "COLORLESS",
    "BEWARE": "REGENT",
    "BLASPHEMY": "COLORLESS",
    "BLIND": "SILENT",
    "BLINK": "DEFECT",
    "BLOCK": "IRONCLAD",
    "BLOODLETTING": "IRONCLAD",
    "BLOOD_PUNCH": "IRONCLAD",
    "BODY_SLAM": "IRONCLAD",
    "BOLT": "DEFECT",
    "BONE": "NECROBINDER",
    "BOOST": "DEFECT",
    "BOTTLE": "REGENT",
    "BRAMBLE": "NECROBINDER",
    "BREACH": "DEFECT",
    "BREAKTHROUGH": "DEFECT",
    "BRUTALITY": "IRONCLAD",
    "BURNING_PACT": "IRONCLAD",
    "CALCIFY": "NECROBINDER",
    "CALTROPS": "COLORLESS",
    "CASCADE": "IRONCLAD",
    "CATALYST": "SILENT",
    "CHAIN": "DEFECT",
    "CHARGED_BOLT": "DEFECT",
    "CHARGE": "REGENT",
    "CHARGE_BATTERY": "DEFECT",
    "CHEAT": "NECROBINDER",
    "CHOP": "SILENT",
    "CLAW": "DEFECT",
    "CLEAR_THE_WAY": "REGENT",
    "CLOTHESLINE": "IRONCLAD",
    "CLUMSY": "COLORLESS",
    "COALESCE": "NECROBINDER",
    "COLD_SNAP": "DEFECT",
    "COMBUST": "IRONCLAD",
    "CONCENTRATE": "SILENT",
    "CONDENSER": "DEFECT",
    "CONDUIT": "DEFECT",
    "CONSECRATE": "IRONCLAD",
    "CONTEMPLATE": "DEFECT",
    "COOLHEAD": "DEFECT",
    "CORE_SURGE": "DEFECT",
    "CORRUPT": "NECROBINDER",
    "COUNTER": "IRONCLAD",
    "CRACKLE": "DEFECT",
    "CRASH": "DEFECT",
    "CRESCENDO": "SILENT",
    "CRUSH": "IRONCLAD",
    "CURSE_OF_THE_BELL": "COLORLESS",
    "CUT": "SILENT",
    "DAGGER_THROW": "SILENT",
    "DARK_EMBRACE": "NECROBINDER",
    "DARK_SHACKLES": "IRONCLAD",
    "DAZED": "NECROBINDER",
    "DEATH_HARROW": "NECROBINDER",
    "DEBILITATE": "NECROBINDER",
    "DEFILE": "NECROBINDER",
    "DEFEND": "IRONCLAD",
    "DEFEND_REGENT": "REGENT",
    "DEFEND_NECROBINDER": "NECROBINDER",
    "DEFLECT": "SILENT",
    "DEMON_FORM": "IRONCLAD",
    "DEVASTATE": "REGENT",
    "DIRE_DOSIS": "DEFECT",
    "DISARM": "IRONCLAD",
    "DISCOVERY": "COLORLESS",
    "DISTRIBUTE": "DEFECT",
    "DODGE_AND_WEAVE": "SILENT",
    "DRAFT": "REGENT",
    "DRAIN": "NECROBINDER",
    "DRAMATIC_ENTRANCE": "REGENT",
    "DREAM_COMPOSER": "DEFECT",
    "DUALCAST": "DEFECT",
    "DYNAMO": "DEFECT",
    "EATER": "DEFECT",
    "EGG": "REGENT",
    "ELECTRODYNAMICS": "DEFECT",
    "EMANATE": "REGENT",
    "EMBED": "DEFECT",
    "EMPTY_BODY": "SILENT",
    "EMPTY_FIST": "IRONCLAD",
    "ENDURE": "IRONCLAD",
    "ENERGIZE": "DEFECT",
    "ENLIGHTEN": "DEFECT",
    "ENRAGE": "IRONCLAD",
    "ENVOY": "REGENT",
    "EQUILIBRIUM": "DEFECT",
    "EVISCERATE": "SILENT",
    "EVOKE_ORB": "DEFECT",
    "EXE": "DEFECT",
    "EXPERTISE": "DEFECT",
    "EXPLODE": "DEFECT",
    "EXPOSURE": "DEFECT",
    "EXTORT": "NECROBINDER",
    "FEAR": "NECROBINDER",
    "FEED": "IRONCLAD",
    "FEEL_NO_PAIN": "IRONCLAD",
    "FIDGET": "SILENT",
    "FINESSE": "COLORLESS",
    "FINISH": "SILENT",
    "FLAME_BURST": "DEFECT",
    "FLAMETHROWER": "DEFECT",
    "FLAP": "REGENT",
    "FLASK": "COLORLESS",
    "FLECHETTES": "SILENT",
    "FLOAT": "DEFECT",
    "FOCUS": "DEFECT",
    "FOLLY": "COLORLESS",
    "FORGE": "IRONCLAD",
    "FORM": "DEFECT",
    "FORTIFY": "NECROBINDER",
    "FURY": "IRONCLAD",
    "GAIN_STRENGTH": "DEFECT",
    "GASH": "IRONCLAD",
    "GHOSTLY": "SILENT",
    "GLACIER": "DEFECT",
    "GLASS": "DEFECT",
    "GOBLIN": "REGENT",
    "GOLDEN_IDOL": "COLORLESS",
    "GREED": "COLORLESS",
    "GRIP": "DEFECT",
    "GUILLOTINE": "IRONCLAD",
    "GUST": "DEFECT",
    "HAUNT": "NECROBINDER",
    "HEATSINK": "DEFECT",
    "HEAVY_BLADE": "IRONCLAD",
    "HEGEMONY": "REGENT",
    "HELIX": "DEFECT",
    "HEMOKINESIS": "IRONCLAD",
    "HEX": "DEFECT",
    "HIDDEN_CACHE": "COLORLESS",
    "HIGHLIGHT": "DEFECT",
    "HODGIES_CLAW": "DEFECT",
    "HOMICIDE": "NECROBINDER",
    "IMPERVIOUS": "IRONCLAD",
    "IMPROVISED_EXPLOSION": "REGENT",
    "INFERNAL_BLAST": "IRONCLAD",
    "INFLAME": "IRONCLAD",
    "INGEST": "NECROBINDER",
    "INITIALIZING_CHARGE": "DEFECT",
    "INSERTAINER": "REGENT",
    "INSIGHT": "COLORLESS",
    "INTIMIDATE": "IRONCLAD",
    "INVEST": "REGENT",
    "IRON_WAVE": "IRONCLAD",
    "JACK": "REGENT",
    "JUGGERNAUT": "IRONCLAD",
    "KNOCKBACK": "IRONCLAD",
    "LEECH": "NECROBINDER",
    "LEFT_HOOK": "REGENT",
    "LEG_SWEEP": "SILENT",
    "LESSON_LEARNED": "DEFECT",
    "LEVELLER": "REGENT",
    "LIGHTNING": "DEFECT",
    "LIMITED": "DEFECT",
    "LOCK_ON": "DEFECT",
    "LOOP": "DEFECT",
    "MAGNETISM": "DEFECT",
    "MALLEABLE": "DEFECT",
    "MASTER_REALITY": "COLORLESS",
    "MAW": "NECROBINDER",
    "MAYHEM": "NECROBINDER",
    "MEGA_DYNAMIC": "IRONCLAD",
    "METALLIC_SCALE": "IRONCLAD",
    "MIND_BLAST": "COLORLESS",
    "MIND_BOLT": "DEFECT",
    "MIRROR": "COLORLESS",
    "MULTICAST": "DEFECT",
    "MULTIPLY": "DEFECT",
    "MUTATION": "REGENT",
    "NEOW_BLESSING": "COLORLESS",
    "NEURO": "DEFECT",
    "NIGHTMARE": "DEFECT",
    "NOBLE": "REGENT",
    "OBLITERATE": "IRONCLAD",
    "OMEGA": "DEFECT",
    "OMNISCIENCE": "COLORLESS",
    "OUTMANEUVER": "REGENT",
    "PALLBEARER": "NECROBINDER",
    "PANACEA": "COLORLESS",
    "PARASITE": "SILENT",
    "PATH_TO_VIOLENCE": "IRONCLAD",
    "PAYDAY": "COLORLESS",
    "PEACE": "REGENT",
    "PELOTON": "DEFECT",
    "PERFECTED_STRIKE": "IRONCLAD",
    "PERFECTLY_FINE": "REGENT",
    "PHANTOM": "SILENT",
    "PHLEGM": "DEFECT",
    "PICK_POCKET": "SILENT",
    "PIERCING_BROW": "DEFECT",
    "PIERCING_WAIL": "SILENT",
    "PILLAGE": "IRONCLAD",
    "POISON": "SILENT",
    "POMMEL_STRIKE": "IRONCLAD",
    "POWER_THROUGH": "IRONCLAD",
    "PRAETORIAN": "REGENT",
    "PRECISION": "SILENT",
    "PREDATOR": "SILENT",
    "PREP_TIME": "COLORLESS",
    "PREPARED": "SILENT",
    "PRIDE": "IRONCLAD",
    "PRIORITY": "DEFECT",
    "PROSTRATE": "IRONCLAD",
    "PROWL": "SILENT",
    "PSIONIC": "DEFECT",
    "PUMMEL": "IRONCLAD",
    "PURE": "DEFECT",
    "QUAKE": "DEFECT",
    "RAIN": "DEFECT",
    "RAMPAGE": "IRONCLAD",
    "REACH": "DEFECT",
    "REACTIVE": "DEFECT",
    "REBOOT": "DEFECT",
    "RECALL": "DEFECT",
    "RECKLESS": "REGENT",
    "RECUR": "DEFECT",
    "RECYCLE": "DEFECT",
    "REDLINE": "DEFECT",
    "REFLECT": "REGENT",
    "REGRET": "COLORLESS",
    "REINFORCE": "DEFECT",
    "RELEASE": "REGENT",
    "REND": "DEFECT",
    "REST": "IRONCLAD",
    "REVENGE": "IRONCLAD",
    "RITUAL": "NECROBINDER",
    "ROUGH": "REGENT",
    "RUTHLESS": "IRONCLAD",
    "SACRIFICE": "NECROBINDER",
    "SAFEGUARD": "REGENT",
    "SANCTITY": "SILENT",
    "SAND": "DEFECT",
    "SCEPTICISM": "DEFECT",
    "SCOOP": "REGENT",
    "SCRAPE": "SILENT",
    "SEARING_BLOW": "IRONCLAD",
    "SECOND_WIND": "IRONCLAD",
    "SECRET_TECHNIQUE": "COLORLESS",
    "SECRET_WEAPON": "COLORLESS",
    "SEDATION": "DEFECT",
    "SEEKER": "REGENT",
    "SEIZE": "NECROBINDER",
    "SELF_REPAIR": "IRONCLAD",
    "SENTINEL": "IRONCLAD",
    "SETUP": "SILENT",
    "SEVER": "NECROBINDER",
    "SHANK": "SILENT",
    "SHARPEN": "IRONCLAD",
    "SHIV": "SILENT",
    "SHOOT": "REGENT",
    "SHORT_STRIDE": "REGENT",
    "SHRUG_IT_OFF": "IRONCLAD",
    "SHUFFLE": "REGENT",
    "SIGHT": "DEFECT",
    "SIGN": "DEFECT",
    "SKIM": "DEFECT",
    "SKIP": "REGENT",
    "SLEEP": "SILENT",
    "SLICE": "REGENT",
    "SLIMED": "COLORLESS",
    "SLOTH": "REGENT",
    "SMITE": "IRONCLAD",
    "SMOKE_BOMB": "SILENT",
    "SNAP": "NECROBINDER",
    "SNIPER": "DEFECT",
    "SOLAR": "IRONCLAD",
    "SONIC_BURST": "DEFECT",
    "SOUL_STRIKE": "NECROBINDER",
    "SPECULATOR": "REGENT",
    "SPIKED_ARMOR": "IRONCLAD",
    "SPIRAL": "DEFECT",
    "SPIT": "DEFECT",
    "SPOT_WEAKNESS": "IRONCLAD",
    "STACK": "REGENT",
    "STAMPEDE": "IRONCLAD",
    "STAR": "DEFECT",
    "STASIS": "DEFECT",
    "STATIC": "DEFECT",
    "STIM": "REGENT",
    "STIM_PACK": "DEFECT",
    "STOCK": "REGENT",
    "STOLE": "SILENT",
    "STORM": "DEFECT",
    "STRIKE": "IRONCLAD",
    "STRIKE_DEFECT": "DEFECT",
    "STRIKE_NECROBINDER": "NECROBINDER",
    "STRIKE_REGENT": "REGENT",
    "STRIKE_SILENT": "SILENT",
    "SUICIDE": "NECROBINDER",
    "SUNDER": "IRONCLAD",
    "SUPLEX": "IRONCLAD",
    "SUPPRESSION": "DEFECT",
    "SURVEILLANCE": "DEFECT",
    "SWEEP": "DEFECT",
    "SWIFT_STRIKE": "SILENT",
    "SWORD": "REGENT",
    "SYNC": "DEFECT",
    "TAKE": "DEFECT",
    "TANK": "IRONCLAD",
    "TEAM_SLAY": "REGENT",
    "TEAR": "DEFECT",
    "TELECAST": "DEFECT",
    "TEMPEST": "DEFECT",
    "TENDRILS": "NECROBINDER",
    "THANATOS": "NECROBINDER",
    "THUNDER": "DEFECT",
    "TINDER": "IRONCLAD",
    "TONE": "DEFECT",
    "TONGUE": "REGENT",
    "TOXIC": "SILENT",
    "TRIP": "SILENT",
    "TURBO": "DEFECT",
    "TWIN_STRIKE": "IRONCLAD",
    "ULTIMATE": "DEFECT",
    "UNCHARGED": "DEFECT",
    "UNDERHAND": "REGENT",
    "UNDERSTANDING": "DEFECT",
    "UNDO": "COLORLESS",
    "UNLOAD": "IRONCLAD",
    "UPPERCUT": "IRONCLAD",
    "UPROOT": "DEFECT",
    "VAULT": "SILENT",
    "VENGEANCE": "REGENT",
    "VENOM": "SILENT",
    "VOLT_EDGE": "DEFECT",
    "VOLTAIC": "DEFECT",
    "WAIL": "DEFECT",
    "WAR_CRY": "IRONCLAD",
    "WARD": "DEFECT",
    "WARP": "DEFECT",
    "WASH": "DEFECT",
    "WATCHER_STANCE": "REGENT",
    "WEAVE": "SILENT",
    "WHEEL": "DEFECT",
    "WHIRLWIND": "IRONCLAD",
    "WISH": "COLORLESS",
    "WONDER": "REGENT",
    "WORLD": "DEFECT",
    "WRATH": "REGENT",
    "WREATH": "REGENT",
    "WRATH_FORM": "REGENT",
    "ZAP": "DEFECT",
    "ZEN": "REGENT",
}

RELIC_INFO = {
    "AKABEKO": {
        "name": "Akabeko",
        "description": "At the start of each combat, gain 8 Vigor.",
    },
    "ALCHEMICAL_COFFER": {
        "name": "Alchemical Coffer",
        "description": "Upon pickup, gain 4 potion slots filled with random potions.",
    },
    "AMETHYST_AUBERGINE": {
        "name": "Amethyst Aubergine",
        "description": "Enemies drop 10 additional Gold.",
    },
    "ANCHOR": {"name": "Anchor", "description": "Start each combat with 10 Block."},
    "ARCANE_SCROLL": {
        "name": "Arcane Scroll",
        "description": "Upon pickup, obtain a random Rare Card to add to your Deck.",
    },
    "ARCHAIC_TOOTH": {
        "name": "Archaic Tooth",
        "description": "Upon pickup, Transform a starter card with an ancient version.",
    },
    "ART_OF_WAR": {
        "name": "Art of War",
        "description": "If you do not play any Attacks during your turn, gain an additional [energy:1] next turn.",
    },
    "ASTROLABE": {
        "name": "Astrolabe",
        "description": "Upon pickup, Transform 3 cards, then Upgrade them.",
    },
    "BAG_OF_MARBLES": {
        "name": "Bag of Marbles",
        "description": "At the start of each combat, apply 1 Vulnerable to ALL enemies.",
    },
    "BAG_OF_PREPARATION": {
        "name": "Bag of Preparation",
        "description": "At the start of each combat, draw 2 additional cards.",
    },
    "BEATING_REMNANT": {
        "name": "Beating Remnant",
        "description": "You cannot lose more than 20 HP in a single turn.",
    },
    "BEAUTIFUL_BRACELET": {
        "name": "Beautiful Bracelet",
        "description": "Upon pickup, choose 3 cards in your Deck. Enchant them with [purple]Swift[/purple] 3.",
    },
    "BELLOWS": {
        "name": "Bellows",
        "description": "The first Hand you draw each combat is Upgraded.",
    },
    "BELT_BUCKLE": {
        "name": "Belt Buckle",
        "description": "While you have no potions, you have 2 additional Dexterity.",
    },
    "BIG_HAT": {
        "name": "Big Hat",
        "description": "At the start of each combat, add 2 random Ethereal cards into your Hand.",
    },
    "BIG_MUSHROOM": {
        "name": "Big Mushroom",
        "description": "Upon pickup, raise your Max HP by 20. At the start of each combat, draw 2 fewer cards.",
    },
    "BIIIG_HUG": {
        "name": "Biiig Hug",
        "description": "Upon pickup, remove 4 cards from your Deck.",
    },
    "BING_BONG": {
        "name": "Bing Bong",
        "description": "Whenever you add a card to your Deck, add one additional copy.",
    },
    "BLACK_BLOOD": {
        "name": "Black Blood",
        "description": "At the end of combat, heal 12 HP.",
    },
    "BLACK_STAR": {
        "name": "Black Star",
        "description": "Elites drop an additional Relic when defeated.",
    },
    "BLESSED_ANTLER": {
        "name": "Blessed Antler",
        "description": "Gain 1 energy at the start of each turn.",
    },
    "BLOOD_SOAKED_ROSE": {
        "name": "Blood-Soaked Rose",
        "description": "Upon pickup, add 1 Enthralled to your Deck.",
    },
    "BLOOD_VIAL": {
        "name": "Blood Vial",
        "description": "At the start of each combat, heal 2 HP.",
    },
    "BONE_FLUTE": {
        "name": "Bone Flute",
        "description": "Whenever Osty attacks, gain 2 Block.",
    },
    "BONE_TEA": {
        "name": "Bone Tea",
        "description": "At the start of the next combat, Upgrade your starting hand.",
    },
    "BOOKMARK": {
        "name": "Bookmark",
        "description": "At the end of each turn, lower the cost of a random Retained card by 1.",
    },
    "BOOK_OF_FIVE_RINGS": {
        "name": "Book of Five Rings",
        "description": "Every 5 cards you add to your Deck, heal 15 HP.",
    },
    "BOOK_REPAIR_KNIFE": {
        "name": "Book Repair Knife",
        "description": "Whenever a non-Minion enemy dies to Doom, heal 3 HP.",
    },
    "BOOMING_CONCH": {
        "name": "Booming Conch",
        "description": "At the start of Elite combats, draw 2 additional cards.",
    },
    "BOUND_PHYLACTERY": {
        "name": "Bound Phylactery",
        "description": "At the start of your turn, Summon 1.",
    },
    "BOWLER_HAT": {"name": "Bowler Hat", "description": "Gain 20% additional Gold."},
    "BREAD": {
        "name": "Bread",
        "description": "At the start of your first turn, lose 2 energy.",
    },
    "BRILLIANT_SCARF": {
        "name": "Brilliant Scarf",
        "description": "The 5th card you play each turn is free.",
    },
    "BRIMSTONE": {
        "name": "Brimstone",
        "description": "At the start of your turn, gain 2 Strength and ALL enemies gain 1 Strength.",
    },
    "BRONZE_SCALES": {
        "name": "Bronze Scales",
        "description": "Start each combat with 3 Thorns.",
    },
    "BURNING_BLOOD": {
        "name": "Burning Blood",
        "description": "At the end of combat, heal 6 HP.",
    },
    "BURNING_STICKS": {
        "name": "Burning Sticks",
        "description": "The first time each combat you Exhaust a Skill, add a copy of it into your Hand.",
    },
    "BYRDPIP": {
        "name": "Byrdpip",
        "description": "Upon pickup, gain the card Byrd Swoop.",
    },
    "CALLING_BELL": {
        "name": "Calling Bell",
        "description": "Upon pickup, obtain a unique Curse and 3 Relics.",
    },
    "CANDELABRA": {
        "name": "Candelabra",
        "description": "At the start of your 2nd turn, gain 2 energy.",
    },
    "CAPTAINS_WHEEL": {
        "name": "Captain's Wheel",
        "description": "At the start of your 3rd turn, gain 18 Block.",
    },
    "CAULDRON": {
        "name": "Cauldron",
        "description": "Upon pickup, brews 5 random potions.",
    },
    "CENTENNIAL_PUZZLE": {
        "name": "Centennial Puzzle",
        "description": "The first time you lose HP each combat, draw 3 cards.",
    },
    "CHANDELIER": {
        "name": "Chandelier",
        "description": "At the start of your 3rd turn, gain 3 energy.",
    },
    "CHARONS_ASHES": {
        "name": "Charon's Ashes",
        "description": "Whenever you Exhaust a card, deal 3 damage to ALL enemies.",
    },
    "CHEMICAL_X": {
        "name": "Chemical X",
        "description": "The effects of your cost X cards are increased by 2.",
    },
    "CHOICES_PARADOX": {
        "name": "Choices Paradox",
        "description": "At the start of each combat, add 1 of 5 random cards into your Hand.",
    },
    "CIRCLET": {"name": "Circlet", "description": "It's a circlet."},
    "CLAWS": {
        "name": "Claws",
        "description": "Upon pickup, Transform up to 6 cards into Maul.",
    },
    "CLOAK_CLASP": {
        "name": "Cloak Clasp",
        "description": "At the end of your turn, gain 1 Block for each card in your Hand.",
    },
    "CRACKED_CORE": {
        "name": "Cracked Core",
        "description": "At the start of each combat, Channel 1 Lightning.",
    },
    "CROSSBOW": {
        "name": "Crossbow",
        "description": "At the start of your turn, add a random Attack into your Hand.",
    },
    "CURSED_PEARL": {
        "name": "Cursed Pearl",
        "description": "Upon pickup, receive Greed. Gain 333 Gold.",
    },
    "DARKSTONE_PERIAPT": {
        "name": "Darkstone Periapt",
        "description": "Whenever you obtain a Curse, raise your Max HP by 6.",
    },
    "DATA_DISK": {
        "name": "Data Disk",
        "description": "Start each combat with 1 Focus.",
    },
    "DAUGHTER_OF_THE_WIND": {
        "name": "Daughter of the Wind",
        "description": "Whenever you play an Attack, gain 1 Block.",
    },
    "DELICATE_FROND": {
        "name": "Delicate Frond",
        "description": "At the start of each combat, fill all empty potion slots with random potions.",
    },
    "DEMON_TONGUE": {
        "name": "Demon Tongue",
        "description": "The first time you lose HP on your turn, heal HP equal to the amount lost.",
    },
    "DIAMOND_DIADEM": {
        "name": "Diamond Diadem",
        "description": "Whenever you play 2 or fewer cards in a turn, take half damage from enemies.",
    },
    "DINGY_RUG": {
        "name": "Dingy Rug",
        "description": "Card rewards can now contain Colorless cards.",
    },
    "DISTINGUISHED_CAPE": {
        "name": "Distinguished Cape",
        "description": "Upon pickup, lose 9 Max HP. Add 3 Apparitions to your Deck.",
    },
    "DIVINE_DESTINY": {
        "name": "Divine Destiny",
        "description": "At the start of each combat, gain 6 Star.",
    },
    "DIVINE_RIGHT": {
        "name": "Divine Right",
        "description": "At the start of each combat, gain 3 Star.",
    },
    "DOLLYS_MIRROR": {
        "name": "Dolly's Mirror",
        "description": "Upon pickup, obtain an additional copy of a card in your Deck.",
    },
    "DRAGON_FRUIT": {
        "name": "Dragon Fruit",
        "description": "Whenever you gain Gold, raise your Max HP by 1.",
    },
    "DREAM_CATCHER": {
        "name": "Dream Catcher",
        "description": "Whenever you Rest, you may add a card to your Deck.",
    },
    "DRIFTWOOD": {
        "name": "Driftwood",
        "description": "You may reroll each card reward once.",
    },
    "DUSTY_TOME": {
        "name": "Dusty Tome",
        "description": "Upon pickup, obtain an Ancient Card.",
    },
    "ECTOPLASM": {
        "name": "Ectoplasm",
        "description": "You can no longer gain Gold. Gain 1 energy at the start of each turn.",
    },
    "ELECTRIC_SHRYMP": {
        "name": "Electric Shrymp",
        "description": "Upon pickup, Enchant a Skill with Imbued.",
    },
    "EMBER_TEA": {
        "name": "Ember Tea",
        "description": "At the start of the next 5 combats, gain 2 Strength.",
    },
    "EMOTION_CHIP": {
        "name": "Emotion Chip",
        "description": "If you lost HP during the previous turn, trigger the passive ability of all Orbs.",
    },
    "EMPTY_CAGE": {
        "name": "Empty Cage",
        "description": "Upon pickup, remove 2 cards from your Deck.",
    },
    "ETERNAL_FEATHER": {
        "name": "Eternal Feather",
        "description": "For every 5 cards in your Deck, heal 3 HP whenever you enter a Rest Site.",
    },
    "FENCING_MANUAL": {
        "name": "Fencing Manual",
        "description": "At the start of each combat, Forge 10.",
    },
    "FESTIVE_POPPER": {
        "name": "Festive Popper",
        "description": "At the start of each combat, deal 9 damage to ALL enemies.",
    },
    "FIDDLE": {
        "name": "Fiddle",
        "description": "At the start of each turn, draw 2 additional cards.",
    },
    "FORGOTTEN_SOUL": {
        "name": "Forgotten Soul",
        "description": "Whenever you Exhaust a card, deal 1 damage to a random enemy.",
    },
    "FRAGRANT_MUSHROOM": {
        "name": "Fragrant Mushroom",
        "description": "Upon pickup, lose 15 HP and Upgrade 3 random cards.",
    },
    "FRESNEL_LENS": {
        "name": "Fresnel Lens",
        "description": "Whenever you add a card that gains Block to your Deck, Enchant it with Nimble 2.",
    },
    "FROZEN_EGG": {
        "name": "Frozen Egg",
        "description": "Whenever you add a Power into your Deck, Upgrade it.",
    },
    "FUNERARY_MASK": {
        "name": "Funerary Mask",
        "description": "At the start of each combat, add 3 Souls into your Draw Pile.",
    },
    "FUR_COAT": {
        "name": "Fur Coat",
        "description": "Upon pickup, mark 7 random combats.",
    },
    "GALACTIC_DUST": {
        "name": "Galactic Dust",
        "description": "For every 10 Star spent, gain 10 Block.",
    },
    "GAMBLING_CHIP": {
        "name": "Gambling Chip",
        "description": "At the start of each combat, discard any number of cards then draw that many.",
    },
    "GAME_PIECE": {
        "name": "Game Piece",
        "description": "Whenever you play a Power, draw 1 card.",
    },
    "GHOST_SEED": {
        "name": "Ghost Seed",
        "description": "Strikes and Defends gain Ethereal.",
    },
    "GIRYA": {
        "name": "Girya",
        "description": "You can now gain Strength at Rest Sites. (3 times max)",
    },
    "GLASS_EYE": {
        "name": "Glass Eye",
        "description": "Upon pickup, obtain 2 Common cards, 2 Uncommon cards, and 1 Rare card.",
    },
    "GLITTER": {
        "name": "Glitter",
        "description": "Enchant all card rewards with Glam.",
    },
    "GNARLED_HAMMER": {
        "name": "Gnarled Hammer",
        "description": "Upon pickup, Enchant up to 3 Attacks with Sharp 3.",
    },
    "GOLDEN_COMPASS": {
        "name": "Golden Compass",
        "description": "Upon pickup, replace the Act 2 Map with a single special path.",
    },
    "GOLDEN_PEARL": {
        "name": "Golden Pearl",
        "description": "Upon pickup, gain 150 Gold.",
    },
    "GOLD_PLATED_CABLES": {
        "name": "Gold-Plated Cables",
        "description": "Your rightmost Orb triggers its passive an additional time.",
    },
    "GORGET": {
        "name": "Gorget",
        "description": "At the start of each combat, gain 4 Plating.",
    },
    "GREMLIN_HORN": {
        "name": "Gremlin Horn",
        "description": "Whenever an enemy dies, gain 1 energy and draw 1 card.",
    },
    "HAND_DRILL": {
        "name": "Hand Drill",
        "description": "Whenever you break an enemy's Block, apply 2 Vulnerable.",
    },
    "HAPPY_FLOWER": {
        "name": "Happy Flower",
        "description": "Every 3 turns, gain 1 energy.",
    },
    "HELICAL_DART": {
        "name": "Helical Dart",
        "description": "Whenever you play a Shiv, gain 1 Dexterity this turn.",
    },
    "HISTORY_COURSE": {
        "name": "History Course",
        "description": "At the start of your turn, play a copy of your last played Attack or Skill.",
    },
    "HORN_CLEAT": {
        "name": "Horn Cleat",
        "description": "At the start of your 2nd turn, gain 14 Block.",
    },
    "ICE_CREAM": {
        "name": "Ice Cream",
        "description": "Energy is now conserved between turns.",
    },
    "INFUSED_CORE": {
        "name": "Infused Core",
        "description": "At the start of each combat, Channel 3 Lightning.",
    },
    "INTIMIDATING_HELMET": {
        "name": "Intimidating Helmet",
        "description": "Whenever you play a card that costs 2 or more, gain 4 Block.",
    },
    "IRON_CLUB": {
        "name": "Iron Club",
        "description": "Every 4 cards you play, draw 1 card.",
    },
    "IVORY_TILE": {
        "name": "Ivory Tile",
        "description": "Whenever you play a card that costs 3 or more, gain 1 energy.",
    },
    "JEWELED_MASK": {
        "name": "Jeweled Mask",
        "description": "At the start of combat put a random Power from your Draw Pile into your Hand.",
    },
    "JEWELRY_BOX": {
        "name": "Jewelry Box",
        "description": "Upon pickup, add 1 Apotheosis to your Deck.",
    },
    "JOSS_PAPER": {
        "name": "Joss Paper",
        "description": "Every 5 times you Exhaust a card, draw 1 card.",
    },
    "JUZU_BRACELET": {
        "name": "Juzu Bracelet",
        "description": "Regular enemy combats are no longer encountered in ? rooms.",
    },
    "KIFUDA": {
        "name": "Kifuda",
        "description": "Upon pickup, Enchant up to 3 cards with Adroit.",
    },
    "KUNAI": {
        "name": "Kunai",
        "description": "Every time you play 3 Attacks in a single turn, gain 1 Dexterity.",
    },
    "KUSARIGAMA": {
        "name": "Kusarigama",
        "description": "Every time you play 3 Attacks in a single turn, deal 6 damage to a random enemy.",
    },
    "LANTERN": {
        "name": "Lantern",
        "description": "Start each combat with an additional 1 energy.",
    },
    "LARGE_CAPSULE": {
        "name": "Large Capsule",
        "description": "Upon pickup, obtain 2 random Relics.",
    },
    "LASTING_CANDY": {
        "name": "Lasting Candy",
        "description": "Every other combat, your card rewards gain an additional Power.",
    },
    "LAVA_LAMP": {
        "name": "Lava Lamp",
        "description": "At the end of combat, Upgrade all card rewards if you took no damage.",
    },
    "LAVA_ROCK": {"name": "Lava Rock", "description": "The Act 1 Boss drops 2 Relics."},
    "LEAD_PAPERWEIGHT": {
        "name": "Lead Paperweight",
        "description": "Upon pickup, choose 1 of 2 Colorless cards to add to your Deck.",
    },
    "LEAFY_POULTICE": {
        "name": "Leafy Poultice",
        "description": "Upon pickup, Transform 1 of your Strikes and 1 of your Defends.",
    },
    "LEES_WAFFLE": {
        "name": "Lee's Waffle",
        "description": "Upon pickup, raise your Max HP by 7 and heal all of your HP.",
    },
    "LETTER_OPENER": {
        "name": "Letter Opener",
        "description": "Every time you play 3 Skills in a single turn, deal 5 damage to ALL enemies.",
    },
    "LIZARD_TAIL": {
        "name": "Lizard Tail",
        "description": "When you would die, heal to 50% of your Max HP instead.",
    },
    "LOOMING_FRUIT": {
        "name": "Looming Fruit",
        "description": "Upon pickup, raise your Max HP by 31.",
    },
    "LORDS_PARASOL": {
        "name": "Lord's Parasol",
        "description": "When you encounter the Merchant, immediately obtain EVERYTHING he sells.",
    },
    "LOST_COFFER": {
        "name": "Lost Coffer",
        "description": "Upon pickup, gain 1 card reward and procure 1 random potion.",
    },
    "LOST_WISP": {
        "name": "Lost Wisp",
        "description": "Whenever you play a Power, deal 8 damage to ALL enemies.",
    },
    "LUCKY_FYSH": {
        "name": "Lucky Fysh",
        "description": "Whenever you add a card to your Deck, gain 15 Gold.",
    },
    "LUNAR_PASTRY": {
        "name": "Lunar Pastry",
        "description": "At the end of your turn, gain 1 Star.",
    },
    "MANGO": {"name": "Mango", "description": "Upon pickup, raise your Max HP by 14."},
    "MASSIVE_SCROLL": {
        "name": "Massive Scroll",
        "description": "Upon pickup, choose 1 of 3 Multiplayer Cards to add to your Deck.",
    },
    "MAW_BANK": {
        "name": "Maw Bank",
        "description": "Whenever you climb a floor, gain 12 Gold.",
    },
    "MEAL_TICKET": {
        "name": "Meal Ticket",
        "description": "Whenever you enter a shop room, heal 15 HP.",
    },
    "MEAT_CLEAVER": {
        "name": "Meat Cleaver",
        "description": "You may Cook at Rest Sites.",
    },
    "MEAT_ON_THE_BONE": {
        "name": "Meat on the Bone",
        "description": "If your HP is at or below 50% at the end of combat, heal 12 HP.",
    },
    "MEMBERSHIP_CARD": {
        "name": "Membership Card",
        "description": "50% discount on all products!",
    },
    "MERCURY_HOURGLASS": {
        "name": "Mercury Hourglass",
        "description": "At the start of your turn, deal 3 damage to ALL enemies.",
    },
    "METRONOME": {
        "name": "Metronome",
        "description": "The first time you Channel 7 Orbs each combat, deal 30 damage to ALL enemies.",
    },
    "MINI_REGENT": {
        "name": "Mini Regent",
        "description": "The first time you spend Star each turn, gain 1 Strength.",
    },
    "MINIATURE_CANNON": {
        "name": "Miniature Cannon",
        "description": "Upgraded Attacks deal 3 additional damage.",
    },
    "MINIATURE_TENT": {
        "name": "Miniature Tent",
        "description": "You may choose any number of options at Rest Sites.",
    },
    "MOLTEN_EGG": {
        "name": "Molten Egg",
        "description": "Whenever you add an Attack card to your Deck, Upgrade it.",
    },
    "MR_STRUGGLES": {
        "name": "Mr. Struggles",
        "description": "At the start of your turn, deal damage equal to the turn number to ALL enemies.",
    },
    "MUMMIFIED_HAND": {
        "name": "Mummified Hand",
        "description": "Whenever you play a Power, a random card in your Hand is free to play that turn.",
    },
    "MUSIC_BOX": {
        "name": "Music Box",
        "description": "Create an Ethereal copy of the first Attack you play each turn.",
    },
    "MYSTIC_LIGHTER": {
        "name": "Mystic Lighter",
        "description": "Enchanted Attacks deal 9 additional damage.",
    },
    "NEOWS_TORMENT": {
        "name": "Neow's Torment",
        "description": "Upon pickup, add 1 Neow's Fury to your Deck.",
    },
    "NEW_LEAF": {"name": "New Leaf", "description": "Upon pickup, Transform 1 card."},
    "NINJA_SCROLL": {
        "name": "Ninja Scroll",
        "description": "At the start of each combat, add 3 Shivs into your Hand.",
    },
    "NUNCHAKU": {
        "name": "Nunchaku",
        "description": "Every time you play 10 Attacks, gain 1 energy.",
    },
    "NUTRITIOUS_OYSTER": {
        "name": "Nutritious Oyster",
        "description": "Upon pickup, raise your Max HP by 11.",
    },
    "NUTRITIOUS_SOUP": {
        "name": "Nutritious Soup",
        "description": "Upon pickup, Enchant all Strikes in your Deck.",
    },
    "ODDLY_SMOOTH_STONE": {
        "name": "Oddly Smooth Stone",
        "description": "Start each combat with 1 Dexterity.",
    },
    "OLD_COIN": {"name": "Old Coin", "description": "Upon pickup, gain 300 Gold."},
    "ORANGE_DOUGH": {
        "name": "Orange Dough",
        "description": "At the start of each combat, add 2 random Colorless cards into your Hand.",
    },
    "ORICHALCUM": {
        "name": "Orichalcum",
        "description": "If you end your turn without Block, gain 6 Block.",
    },
    "ORNAMENTAL_FAN": {
        "name": "Ornamental Fan",
        "description": "Every time you play 3 Attacks in a single turn, gain 4 Block.",
    },
    "ORRERY": {"name": "Orrery", "description": "Upon pickup, gain 5 card rewards."},
    "PAELS_BLOOD": {
        "name": "Pael's Blood",
        "description": "At the start of your turn, draw 1 additional card.",
    },
    "PAELS_CLAW": {
        "name": "Pael's Claw",
        "description": "Upon pickup, Enchant all Defends with Goopy.",
    },
    "PAELS_EYE": {
        "name": "Pael's Eye",
        "description": "The first time each combat you end your turn without playing cards.",
    },
    "PAELS_FLESH": {
        "name": "Pael's Flesh",
        "description": "Gain an additional energy at the start of your 3rd turn.",
    },
    "PAELS_GROWTH": {
        "name": "Pael's Growth",
        "description": "Upon pickup, Enchant a card with Clone.",
    },
    "PAELS_HORN": {
        "name": "Pael's Horn",
        "description": "Upon pickup, add 2 Relax to your Deck.",
    },
    "PAELS_LEGION": {
        "name": "Pael's Legion",
        "description": "Doubles Block gained from a card, then goes to sleep for 2 turns.",
    },
    "PAELS_TEARS": {
        "name": "Pael's Tears",
        "description": "If you end your turn with unspent energy, gain additional energy next turn.",
    },
    "PAELS_TOOTH": {
        "name": "Pael's Tooth",
        "description": "Upon pickup, remove 5 cards from your Deck.",
    },
    "PAELS_WING": {
        "name": "Pael's Wing",
        "description": "You may sacrifice card rewards to Pael.",
    },
    "PANDORAS_BOX": {
        "name": "Pandora's Box",
        "description": "Transform ALL Strikes and Defends.",
    },
    "PANTOGRAPH": {
        "name": "Pantograph",
        "description": "At the start of each Boss combat, heal 25 HP.",
    },
    "PAPER_KRANE": {
        "name": "Paper Krane",
        "description": "Enemies with Weak deal 40% less damage to you.",
    },
    "PAPER_PHROG": {
        "name": "Paper Phrog",
        "description": "Enemies with Vulnerable take 75% more damage.",
    },
    "PARRYING_SHIELD": {
        "name": "Parrying Shield",
        "description": "If you end a turn with at least 10 Block, deal 6 damage to a random enemy.",
    },
    "PEAR": {"name": "Pear", "description": "Upon pickup, raise your Max HP by 10."},
    "PEN_NIB": {
        "name": "Pen Nib",
        "description": "Every 10th Attack you play deals double damage.",
    },
    "PENDULUM": {
        "name": "Pendulum",
        "description": "Whenever you shuffle your Draw Pile, draw a card.",
    },
    "PERMAFROST": {
        "name": "Permafrost",
        "description": "The first time you play a Power each combat, gain 6 Block.",
    },
    "PETRIFIED_TOAD": {
        "name": "Petrified Toad",
        "description": "At the start of each combat, procure a Potion-Shaped Rock.",
    },
    "PHILOSOPHERS_STONE": {
        "name": "Philosopher's Stone",
        "description": "Gain 1 energy at the start of each turn. ALL enemies start combat with 1 Strength.",
    },
    "PHYLACTERY_UNBOUND": {
        "name": "Phylactery Unbound",
        "description": "At the start of each combat, Summon 5. At the start of your turn, Summon 2.",
    },
    "PLANISPHERE": {
        "name": "Planisphere",
        "description": "Whenever you enter a ? room, heal 4 HP.",
    },
    "POCKETWATCH": {
        "name": "Pocketwatch",
        "description": "Whenever you play 3 or fewer cards during your turn, draw 3 additional cards.",
    },
    "POLLINOUS_CORE": {
        "name": "Pollinous Core",
        "description": "Every 4 turns, draw 2 additional cards.",
    },
    "POMANDER": {"name": "Pomander", "description": "Upon pickup, Upgrade a card."},
    "POTION_BELT": {
        "name": "Potion Belt",
        "description": "Upon pickup, gain 2 potion slots.",
    },
    "POWER_CELL": {
        "name": "Power Cell",
        "description": "At the start of each combat, add 2 zero-cost cards from your Draw Pile into your Hand.",
    },
    "PRAYER_WHEEL": {
        "name": "Prayer Wheel",
        "description": "Normal enemies drop an additional card reward.",
    },
    "PRECARIOUS_SHEARS": {
        "name": "Precarious Shears",
        "description": "Upon pickup, remove 2 cards from your Deck.",
    },
    "PRECISE_SCISSORS": {
        "name": "Precise Scissors",
        "description": "Upon pickup, remove 1 card from your Deck.",
    },
    "PRESERVED_FOG": {
        "name": "Preserved Fog",
        "description": "Upon pickup, remove 5 cards from your Deck.",
    },
    "PRISMATIC_GEM": {
        "name": "Prismatic Gem",
        "description": "Gain 1 energy at the start of each turn. Card rewards contain cards from other colors.",
    },
    "PUMPKIN_CANDLE": {
        "name": "Pumpkin Candle",
        "description": "Gain 1 energy at the start of each turn.",
    },
    "PUNCH_DAGGER": {
        "name": "Punch Dagger",
        "description": "Upon pickup, Enchant an Attack with Momentum 5.",
    },
    "RADIANT_PEARL": {
        "name": "Radiant Pearl",
        "description": "At the start of each combat, add 1 Luminesce into your Hand.",
    },
    "RAINBOW_RING": {
        "name": "Rainbow Ring",
        "description": "The first time you play an Attack, Skill, and Power each turn, gain 1 Strength and 1 Dexterity.",
    },
    "RAZOR_TOOTH": {
        "name": "Razor Tooth",
        "description": "Every time you play an Attack or Skill, Upgrade it for the remainder of combat.",
    },
    "RED_MASK": {
        "name": "Red Mask",
        "description": "At the start of each combat, apply 1 Weak to ALL enemies.",
    },
    "RED_SKULL": {
        "name": "Red Skull",
        "description": "While your HP is at or below 50%, you have 3 additional Strength.",
    },
    "REGAL_PILLOW": {
        "name": "Regal Pillow",
        "description": "Whenever you Rest, heal an additional 15 HP.",
    },
    "REGALITE": {
        "name": "Regalite",
        "description": "Whenever you create a Colorless card, gain 2 Block.",
    },
    "REPTILE_TRINKET": {
        "name": "Reptile Trinket",
        "description": "Whenever you use a potion, gain 3 Strength this turn.",
    },
    "RINGING_TRIANGLE": {
        "name": "Ringing Triangle",
        "description": "Retain your Hand on the first turn of combat.",
    },
    "RING_OF_THE_DRAKE": {
        "name": "Ring of the Drake",
        "description": "At the start of your first 3 turns, draw 2 additional cards.",
    },
    "RING_OF_THE_SNAKE": {
        "name": "Ring of the Snake",
        "description": "At the start of each combat, draw 2 additional cards.",
    },
    "RIPPLE_BASIN": {
        "name": "Ripple Basin",
        "description": "If you did not play any Attacks during your turn, gain 4 Block.",
    },
    "ROYAL_POISON": {
        "name": "Royal Poison",
        "description": "At the start of each combat, lose 4 HP.",
    },
    "ROYAL_STAMP": {
        "name": "Royal Stamp",
        "description": "Upon pickup, choose an Attack or Skill in your Deck to Enchant.",
    },
    "RUINED_HELMET": {
        "name": "Ruined Helmet",
        "description": "The first time you gain Strength each combat, double the amount gained.",
    },
    "RUNIC_CAPACITOR": {
        "name": "Runic Capacitor",
        "description": "Start each combat with 3 additional Orb Slots.",
    },
    "RUNIC_PYRAMID": {
        "name": "Runic Pyramid",
        "description": "At the end of your turn, you no longer discard your Hand.",
    },
    "SAI": {"name": "Sai", "description": "At the start of your turn, gain 7 Block."},
    "SAND_CASTLE": {
        "name": "Sand Castle",
        "description": "Upon pickup, Upgrade 6 random cards.",
    },
    "SCREAMING_FLAGON": {
        "name": "Screaming Flagon",
        "description": "If you end your turn with no cards in your Hand, deal 20 damage to ALL enemies.",
    },
    "SCROLL_BOXES": {
        "name": "Scroll Boxes",
        "description": "Upon pickup, lose all Gold and choose 1 of 2 packs of cards.",
    },
    "SEA_GLASS": {
        "name": "Sea Glass",
        "description": "See 15 cards from another character.",
    },
    "SEAL_OF_GOLD": {
        "name": "Seal of Gold",
        "description": "At the start of your turn, spend 5 Gold to gain 1 energy.",
    },
    "SELF_FORMING_CLAY": {
        "name": "Self-Forming Clay",
        "description": "Whenever you lose HP in combat, gain 3 Block next turn.",
    },
    "SERE_TALON": {
        "name": "Sere Talon",
        "description": "Upon pickup, add 2 random Curses and 3 Wishes to your Deck.",
    },
    "SHOVEL": {
        "name": "Shovel",
        "description": "You can now dig at Rest Sites to obtain a random Relic.",
    },
    "SHURIKEN": {
        "name": "Shuriken",
        "description": "Every time you play 3 Attacks in a single turn, gain 1 Strength.",
    },
    "SIGNET_RING": {
        "name": "Signet Ring",
        "description": "Upon pickup, gain 999 Gold.",
    },
    "SILVER_CRUCIBLE": {
        "name": "Silver Crucible",
        "description": "The first 3 card rewards you see are Upgraded.",
    },
    "SLING_OF_COURAGE": {
        "name": "Sling of Courage",
        "description": "Start each Elite combat with 2 Strength.",
    },
    "SMALL_CAPSULE": {
        "name": "Small Capsule",
        "description": "Upon pickup, obtain a random Relic.",
    },
    "SNECKO_EYE": {
        "name": "Snecko Eye",
        "description": "At the start of your turn, draw 2 additional cards.",
    },
    "SNECKO_SKULL": {
        "name": "Snecko Skull",
        "description": "Whenever you apply Poison, apply an additional 1 Poison.",
    },
    "SOZU": {
        "name": "Sozu",
        "description": "Gain 1 energy at the start of each turn. You can no longer obtain potions.",
    },
    "SPARKLING_ROUGE": {
        "name": "Sparkling Rouge",
        "description": "At the start of your 3rd turn, gain 1 Strength and 1 Dexterity.",
    },
    "SPIKED_GAUNTLETS": {
        "name": "Spiked Gauntlets",
        "description": "Gain 1 energy at the start of each turn. Powers cost 1 more.",
    },
    "STONE_CALENDAR": {
        "name": "Stone Calendar",
        "description": "At the end of turn 7, deal 52 damage to ALL enemies.",
    },
    "STONE_CRACKER": {
        "name": "Stone Cracker",
        "description": "At the start of Boss combats, Upgrade 3 random cards.",
    },
    "STONE_HUMIDIFIER": {
        "name": "Stone Humidifier",
        "description": "Whenever you Rest at a Rest Site, raise your Max HP by 5.",
    },
    "STORYBOOK": {
        "name": "Storybook",
        "description": "Upon pickup, add 1 Brightest Flame to your Deck.",
    },
    "STRAWBERRY": {
        "name": "Strawberry",
        "description": "Upon pickup, raise your Max HP by 7.",
    },
    "STRIKE_DUMMY": {
        "name": "Strike Dummy",
        "description": "Cards containing Strike deal 3 additional damage.",
    },
    "STURDY_CLAMP": {
        "name": "Sturdy Clamp",
        "description": "Up to 10 Block persists across turns.",
    },
    "SWORD_OF_JADE": {
        "name": "Sword of Jade",
        "description": "Start each combat with 3 Strength.",
    },
    "SWORD_OF_STONE": {
        "name": "Sword of Stone",
        "description": "Transforms into a powerful Relic after defeating 5 Elites.",
    },
    "SYMBIOTIC_VIRUS": {
        "name": "Symbiotic Virus",
        "description": "At the start of each combat, Channel 1 Dark.",
    },
    "TANXS_WHISTLE": {
        "name": "Tanx's Whistle",
        "description": "Upon pickup, add 1 Whistle to your Deck.",
    },
    "TEA_OF_DISCOURTESY": {
        "name": "Tea of Discourtesy",
        "description": "At the start of the next combat, shuffle 2 Dazed into your Draw Pile.",
    },
    "THE_ABACUS": {
        "name": "The Abacus",
        "description": "Whenever you shuffle your Draw Pile, gain 6 Block.",
    },
    "THE_BOOT": {
        "name": "The Boot",
        "description": "Whenever you would deal 4 or less unblocked attack damage, increase it to 5.",
    },
    "THE_COURIER": {
        "name": "The Courier",
        "description": "The merchant no longer runs out of cards, relics, or potions.",
    },
    "THROWING_AXE": {
        "name": "Throwing Axe",
        "description": "The first card you play each combat is played an extra time.",
    },
    "TINGSHA": {
        "name": "Tingsha",
        "description": "Whenever you discard a card during your turn, deal 3 damage to a random enemy.",
    },
    "TINY_MAILBOX": {
        "name": "Tiny Mailbox",
        "description": "Whenever you Rest, procure a random potion.",
    },
    "TOASTY_MITTENS": {
        "name": "Toasty Mittens",
        "description": "At the start of your turn, Exhaust the top card of your Draw Pile and gain 1 Strength.",
    },
    "TOOLBOX": {
        "name": "Toolbox",
        "description": "At the start of each combat, choose 1 of 3 random Colorless cards.",
    },
    "TOUCH_OF_OROBAS": {
        "name": "Touch of Orobas",
        "description": "Upon pickup, replace your starter Relic with an Ancient version.",
    },
    "TOUGH_BANDAGES": {
        "name": "Tough Bandages",
        "description": "Whenever you discard a card during your turn, gain 3 Block.",
    },
    "TOXIC_EGG": {
        "name": "Toxic Egg",
        "description": "Whenever you add a Skill into your Deck, Upgrade it.",
    },
    "TOY_BOX": {"name": "Toy Box", "description": "Upon pickup, obtain 4 Wax Relics."},
    "TRI_BOOMERANG": {
        "name": "Tri-Boomerang",
        "description": "Choose 3 Attacks in your Deck. Enchant them with Instinct.",
    },
    "TUNGSTEN_ROD": {
        "name": "Tungsten Rod",
        "description": "Whenever you would lose HP, lose 1 less.",
    },
    "TUNING_FORK": {
        "name": "Tuning Fork",
        "description": "Every time you play 10 Skills, gain 7 Block.",
    },
    "TWISTED_FUNNEL": {
        "name": "Twisted Funnel",
        "description": "At the start of each combat, apply 4 Poison to ALL enemies.",
    },
    "UNCEASING_TOP": {
        "name": "Unceasing Top",
        "description": "Whenever you have no cards in Hand during your turn, draw a card.",
    },
    "UNDYING_SIGIL": {
        "name": "Undying Sigil",
        "description": "Enemies with at least as much Doom as HP deal 50% less damage.",
    },
    "UNSETTLING_LAMP": {
        "name": "Unsettling Lamp",
        "description": "Each combat, the first time you play a card that Debuffs an enemy, double its effect.",
    },
    "VAJRA": {"name": "Vajra", "description": "Start each combat with 1 Strength."},
    "VAMBRACE": {
        "name": "Vambrace",
        "description": "The first time you gain Block from a card each combat, double the amount gained.",
    },
    "VELVET_CHOKER": {
        "name": "Velvet Choker",
        "description": "Gain 1 energy at the start of each turn. You cannot play more than 6 cards per turn.",
    },
    "VENERABLE_TEA_SET": {
        "name": "Venerable Tea Set",
        "description": "Whenever you enter a Rest Site, start the next combat with an additional 2 energy.",
    },
    "VERY_HOT_COCOA": {
        "name": "Very Hot Cocoa",
        "description": "Start each combat with an additional 4 energy.",
    },
    "VEXING_PUZZLEBOX": {
        "name": "Vexing Puzzlebox",
        "description": "At the start of each combat, add a random card into your Hand.",
    },
    "VITRUVIAN_MINION": {
        "name": "Vitruvian Minion",
        "description": "Cards containing Minion deal double damage.",
    },
    "WAR_HAMMER": {
        "name": "War Hammer",
        "description": "Whenever you kill an Elite, Upgrade 4 random cards.",
    },
    "WAR_PAINT": {
        "name": "War Paint",
        "description": "Upon pickup, Upgrade 2 random Skills.",
    },
    "WHETSTONE": {
        "name": "Whetstone",
        "description": "Upon pickup, Upgrade 2 random Attacks.",
    },
    "WHISPERING_EARRING": {
        "name": "Whispering Earring",
        "description": "Gain 1 energy at the start of each turn.",
    },
    "WHITE_BEAST_STATUE": {
        "name": "White Beast Statue",
        "description": "Potions always appear in combat rewards.",
    },
    "WHITE_STAR": {
        "name": "White Star",
        "description": "Elites drop an additional Rare card reward.",
    },
    "WING_CHARM": {
        "name": "Wing Charm",
        "description": "A random card in each card reward is Enchanted with Swift 1.",
    },
    "YUMMY_COOKIE": {
        "name": "Yummy Cookie",
        "description": "Upon pickup, Upgrade 4 cards.",
    },
}


def find_sts2_history_dir():
    if sys.platform == "win32":
        paths = [
            os.path.join(os.environ.get("APPDATA", ""), "SlayTheSpire2", "steam"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "SlayTheSpire2", "steam"),
        ]
    elif sys.platform == "darwin":
        paths = [
            os.path.expanduser("~/Library/Application Support/SlayTheSpire2/steam/")
        ]
    else:
        paths = [os.path.expanduser("~/.local/share/SlayTheSpire2/steam/")]

    for base_path in paths:
        if not os.path.exists(base_path):
            continue
        try:
            for entry in os.listdir(base_path):
                profile_base = os.path.join(base_path, entry)
                for subpath in ["profile1/saves/history", "saves/history"]:
                    profile_path = os.path.join(profile_base, subpath)
                    if os.path.isdir(profile_path) and glob.glob(
                        os.path.join(profile_path, "*.run")
                    ):
                        return profile_path
        except PermissionError:
            continue
    return None


def load_runs(history_dir):
    runs = []
    for f in sorted(glob.glob(os.path.join(history_dir, "*.run"))):
        try:
            with open(f, "r") as file:
                runs.append(json.load(file))
        except:
            continue
    return runs


def get_card_data(runs):
    pick_by_class = defaultdict(
        lambda: defaultdict(lambda: {"offered": 0, "picked": 0})
    )
    win_by_class = defaultdict(lambda: defaultdict(lambda: {"runs": 0, "wins": 0}))

    for run in runs:
        is_win = run.get("win", False)
        players = run.get("players", [])
        if not players:
            continue
        char = players[0].get("character", "Unknown").replace("CHARACTER.", "")
        if char not in CLASS_COLORS:
            continue

        cards_in_run = set()
        for act in run.get("map_point_history", []):
            for point in act:
                for stat in point.get("player_stats", []):
                    for choice in stat.get("card_choices", []):
                        card_id = (
                            choice.get("card", {})
                            .get("id", "Unknown")
                            .replace("CARD.", "")
                        )
                        was_picked = choice.get("was_picked", False)
                        pick_by_class[char][card_id]["offered"] += 1
                        if was_picked:
                            pick_by_class[char][card_id]["picked"] += 1
                            cards_in_run.add(card_id)

        for cid in cards_in_run:
            win_by_class[char][cid]["runs"] += 1
            if is_win:
                win_by_class[char][cid]["wins"] += 1

    return pick_by_class, win_by_class


def get_relic_data(runs):
    relic_stats = defaultdict(lambda: {"total": 0, "wins": 0})
    for run in runs:
        is_win = run.get("win", False)
        players = run.get("players", [])
        if not players:
            continue
        for relic in players[0].get("relics", []):
            rid = relic.get("id", "Unknown").replace("RELIC.", "")
            relic_stats[rid]["total"] += 1
            if is_win:
                relic_stats[rid]["wins"] += 1
    return relic_stats


def create_excel(pick_by_class, win_by_class, relic_stats, output_path):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for class_name in CLASS_COLORS.keys():
        ws = wb.create_sheet(title=class_name)
        header_fill = PatternFill(
            start_color="FF" + CLASS_COLORS[class_name],
            end_color="FF" + CLASS_COLORS[class_name],
            fill_type="solid",
        )
        headers = [
            "Card",
            "Original Class",
            "Times Offered",
            "Times Picked",
            "Pick Rate %",
            "Runs with Card",
            "Wins with Card",
            "Win Rate %",
        ]

        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

        row = 2
        class_cards = list(pick_by_class[class_name].items())
        class_cards.sort(key=lambda x: x[1]["picked"], reverse=True)

        for card_id, pick_stats in class_cards:
            win_stats = win_by_class[class_name].get(card_id, {"runs": 0, "wins": 0})
            original_class = CARD_CLASSIFICATIONS.get(card_id, "COLORLESS")
            pick_rate = (
                (pick_stats["picked"] / pick_stats["offered"] * 100)
                if pick_stats["offered"] > 0
                else 0
            )
            win_rate = (
                (win_stats["wins"] / win_stats["runs"] * 100)
                if win_stats["runs"] > 0
                else 0
            )

            ws.cell(row=row, column=1, value=card_id)
            ws.cell(row=row, column=2, value=original_class)
            ws.cell(row=row, column=3, value=pick_stats["offered"])
            ws.cell(row=row, column=4, value=pick_stats["picked"])
            ws.cell(row=row, column=5, value=round(pick_rate, 1))
            ws.cell(row=row, column=6, value=win_stats["runs"])
            ws.cell(row=row, column=7, value=win_stats["wins"])
            ws.cell(row=row, column=8, value=round(win_rate, 1))
            row += 1

        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 15
        for col in range(3, 9):
            ws.column_dimensions[get_column_letter(col)].width = 14

    # Add Relics sheet
    ws = wb.create_sheet(title="Relics")
    header_fill = PatternFill(
        start_color="CCCCCC", end_color="CCCCCC", fill_type="solid"
    )
    headers = ["ID", "Name", "Runs", "Wins", "Win%", "", "Description"]

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    row = 2
    relics_list = [(rid, stats) for rid, stats in relic_stats.items()]
    relics_list.sort(key=lambda x: x[1]["total"], reverse=True)

    for relic_id, stats in relics_list:
        win_rate = (stats["wins"] / stats["total"] * 100) if stats["total"] > 0 else 0
        relic_info = RELIC_INFO.get(
            relic_id, {"name": relic_id, "description": "Unknown"}
        )

        ws.cell(row=row, column=1, value=relic_id)
        ws.cell(row=row, column=2, value=relic_info.get("name", relic_id))
        ws.cell(row=row, column=3, value=stats["total"])
        ws.cell(row=row, column=4, value=stats["wins"])
        ws.cell(row=row, column=5, value=round(win_rate, 1))
        ws.cell(row=row, column=6, value="")
        ws.cell(row=row, column=7, value=relic_info.get("description", ""))
        row += 1

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 5
    ws.column_dimensions["G"].width = 55

    wb.save(output_path)
    return sum(len(cards) for cards in pick_by_class.values())


class STSCardViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("STS2 Statistics Viewer")
        self.geometry("1200x700")
        self.history_dir = None
        self.card_data = {}
        self.relic_data = []
        self.all_data = []
        self.current_class = None

        self.create_widgets()
        self.find_and_load_data()

    def create_widgets(self):
        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(top_frame, text="STS2 Save Location:").pack(side="left", padx=5)
        self.path_label = ttk.Label(top_frame, text="Not selected", foreground="gray")
        self.path_label.pack(side="left", padx=5)
        ttk.Button(top_frame, text="Browse...", command=self.browse_folder).pack(
            side="left", padx=5
        )
        ttk.Button(top_frame, text="Generate Data", command=self.generate_data).pack(
            side="left", padx=5
        )
        ttk.Button(top_frame, text="Help", command=self.show_help).pack(
            side="right", padx=5
        )

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        self.cards_frame = ttk.Frame(self.notebook)
        self.relics_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.cards_frame, text="Cards")
        self.notebook.add(self.relics_frame, text="Relics")

        self.create_cards_tab()
        self.create_relics_tab()

        self.status_label = ttk.Label(self, text="", relief="sunken", anchor="w")
        self.status_label.pack(fill="x", padx=5, pady=2)

    def create_cards_tab(self):
        toolbar = ttk.Frame(self.cards_frame)
        toolbar.pack(fill="x", padx=5, pady=5)

        ttk.Label(toolbar, text="Class:").pack(side="left", padx=5)
        self.class_var = tk.StringVar(value="IRONCLAD")
        for class_name in CLASS_COLORS.keys():
            rb = ttk.Radiobutton(
                toolbar,
                text=class_name,
                value=class_name,
                variable=self.class_var,
                command=self.on_class_change,
            )
            rb.pack(side="left", padx=2)

        ttk.Label(toolbar, text="Search:").pack(side="left", padx=(20, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(toolbar, textvariable=self.search_var, width=20)
        self.search_entry.pack(side="left")
        self.search_entry.bind("<KeyRelease>", lambda e: self.filter_data())
        ttk.Button(toolbar, text="Refresh", command=self.refresh).pack(
            side="right", padx=5
        )

        main_frame = ttk.Frame(self.cards_frame)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)

        columns = ("Card", "Orig", "Offered", "Picked", "Pick%", "Runs", "Wins", "Win%")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings")

        self.tree.heading("Card", text="Card", command=lambda: self.sort_by("card"))
        self.tree.heading(
            "Orig", text="Orig Class", command=lambda: self.sort_by("original_class")
        )
        self.tree.heading(
            "Offered", text="Times Offered", command=lambda: self.sort_by("offered")
        )
        self.tree.heading(
            "Picked", text="Times Picked", command=lambda: self.sort_by("picked")
        )
        self.tree.heading(
            "Pick%", text="Pick Rate %", command=lambda: self.sort_by("pick_rate")
        )
        self.tree.heading(
            "Runs", text="Runs with Card", command=lambda: self.sort_by("runs")
        )
        self.tree.heading(
            "Wins", text="Wins with Card", command=lambda: self.sort_by("wins")
        )
        self.tree.heading(
            "Win%", text="Win Rate %", command=lambda: self.sort_by("win_rate")
        )

        self.tree.column("Card", width=180)
        self.tree.column("Orig", width=80, anchor="center")
        for col in columns[2:]:
            self.tree.column(col, width=70, anchor="e")

        scrollbar = ttk.Scrollbar(
            main_frame, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.tree.tag_configure("oddrow", background="#f0f0f0")
        self.tree.tag_configure("evenrow", background="#ffffff")

    def create_relics_tab(self):
        toolbar = ttk.Frame(self.relics_frame)
        toolbar.pack(fill="x", padx=5, pady=5)

        ttk.Label(toolbar, text="Search:").pack(side="left", padx=5)
        self.relic_search_var = tk.StringVar()
        self.relic_search_entry = ttk.Entry(
            toolbar, textvariable=self.relic_search_var, width=20
        )
        self.relic_search_entry.pack(side="left")
        self.relic_search_entry.bind("<KeyRelease>", lambda e: self.filter_relics())
        ttk.Button(toolbar, text="Refresh", command=self.refresh).pack(
            side="right", padx=5
        )

        main_frame = ttk.Frame(self.relics_frame)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)

        columns = ("Name", "ID", "Runs", "Wins", "Win%", "Description")
        self.relic_tree = ttk.Treeview(
            main_frame, columns=columns, show="headings", selectmode="browse"
        )

        self.relic_tree.heading(
            "Name", text="Relic Name", command=lambda: self.sort_relics("name")
        )
        self.relic_tree.heading(
            "ID", text="Relic ID", command=lambda: self.sort_relics("id")
        )
        self.relic_tree.heading(
            "Runs", text="Runs", command=lambda: self.sort_relics("total")
        )
        self.relic_tree.heading(
            "Wins", text="Wins", command=lambda: self.sort_relics("wins")
        )
        self.relic_tree.heading(
            "Win%", text="Win Rate %", command=lambda: self.sort_relics("win_rate")
        )
        self.relic_tree.heading(
            "Description",
            text="Description",
            command=lambda: self.sort_relics("description"),
        )

        self.relic_tree.column("Name", width=150)
        self.relic_tree.column("ID", width=150)
        self.relic_tree.column("Runs", width=60, anchor="e")
        self.relic_tree.column("Wins", width=60, anchor="e")
        self.relic_tree.column("Win%", width=60, anchor="e")
        self.relic_tree.column("Description", width=500)

        scrollbar_y = ttk.Scrollbar(
            main_frame, orient="vertical", command=self.relic_tree.yview
        )
        scrollbar_x = ttk.Scrollbar(
            main_frame, orient="horizontal", command=self.relic_tree.xview
        )
        self.relic_tree.configure(
            yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set
        )

        self.relic_tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")

        self.relic_tree.tag_configure("oddrow", background="#f0f0f0")
        self.relic_tree.tag_configure("evenrow", background="#ffffff")

        self.relic_desc_label = ttk.Label(
            self.relics_frame, text="", wraplength=700, justify="left"
        )
        self.relic_desc_label.pack(fill="x", padx=5, pady=2)
        self.relic_tree.bind("<<TreeviewSelect>>", self.on_relic_select)

    def find_and_load_data(self):
        self.history_dir = find_sts2_history_dir()
        if self.history_dir:
            self.path_label.config(text=self.history_dir, foreground="black")
            self.generate_data()
        else:
            self.path_label.config(
                text="Not found - click Browse to select", foreground="red"
            )

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select STS2 save folder")
        if folder:
            history_path = os.path.join(folder, "saves", "history")
            if os.path.exists(history_path) and glob.glob(
                os.path.join(history_path, "*.run")
            ):
                self.history_dir = history_path
                self.path_label.config(text=history_path, foreground="black")
                self.generate_data()
            else:
                self.history_dir = folder
                self.path_label.config(text=folder, foreground="black")
                self.generate_data()

    def generate_data(self):
        if not self.history_dir:
            messagebox.showwarning("Warning", "Please select a STS2 save folder first.")
            return

        self.status_label.config(text="Loading runs...")
        self.update()

        runs = load_runs(self.history_dir)
        if not runs:
            messagebox.showwarning(
                "No Data", "No run files found in the selected folder."
            )
            self.status_label.config(text="No runs found")
            return

        self.status_label.config(text=f"Processing {len(runs)} runs...")
        self.update()

        pick_by_class, win_by_class = get_card_data(runs)
        relic_stats = get_relic_data(runs)
        total = create_excel(pick_by_class, win_by_class, relic_stats, EXCEL_FILE)

        self.status_label.config(
            text=f"Generated {total} cards and {len(relic_stats)} relics from {len(runs)} runs"
        )
        self.load_excel_data()

    def load_excel_data(self):
        if not os.path.exists(EXCEL_FILE):
            return

        self.card_data = {}
        self.relic_data = []
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            for class_name in wb.sheetnames:
                ws = wb[class_name]
                if class_name == "Relics":
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if row[0]:
                            self.relic_data.append(
                                {
                                    "id": row[0],
                                    "name": row[1],
                                    "total": row[2],
                                    "wins": row[3],
                                    "win_rate": row[4],
                                    "description": row[6] if len(row) > 6 else "",
                                }
                            )
                else:
                    self.card_data[class_name] = []
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if row[0]:
                            self.card_data[class_name].append(
                                {
                                    "card": row[0],
                                    "original_class": row[1],
                                    "offered": row[2],
                                    "picked": row[3],
                                    "pick_rate": row[4],
                                    "runs": row[5],
                                    "wins": row[6],
                                    "win_rate": row[7],
                                }
                            )
            self.load_class("IRONCLAD")
            self.load_relics()
            runs = load_runs(self.history_dir) if self.history_dir else []
            self.status_label.config(
                text=f"Loaded {sum(len(c) for c in self.card_data.values())} cards and {len(self.relic_data)} relics from {len(runs)} runs"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {e}")

    def load_class(self, class_name):
        self.current_class = class_name
        self.all_data = self.card_data.get(class_name, [])
        self.filter_data()

    def filter_data(self):
        search = self.search_var.get().lower()
        filtered = self.all_data
        if search:
            filtered = [c for c in self.all_data if search in c["card"].lower()]

        for row in self.tree.get_children():
            self.tree.delete(row)

        for i, card in enumerate(filtered):
            tags = ("oddrow",) if i % 2 == 0 else ("evenrow",)
            self.tree.insert(
                "",
                "end",
                values=(
                    card["card"],
                    card["original_class"],
                    card["offered"],
                    card["picked"],
                    card["pick_rate"],
                    card["runs"],
                    card["wins"],
                    card["win_rate"],
                ),
                tags=tags,
            )

    def sort_by(self, key):
        reverse = getattr(self, "_sort_reverse", False)
        self._sort_reverse = not reverse

        def get_val(card):
            val = card.get(key)
            if isinstance(val, str):
                return val.lower()
            return val if val else 0

        self.all_data.sort(key=get_val, reverse=reverse)
        self.filter_data()

    def on_class_change(self):
        self.load_class(self.class_var.get())

    def load_relics(self):
        self.filter_relics()

    def filter_relics(self):
        search = self.relic_search_var.get().lower()
        filtered = self.relic_data
        if search:
            filtered = [
                r
                for r in self.relic_data
                if search in r["name"].lower() or search in r["id"].lower()
            ]

        for row in self.relic_tree.get_children():
            self.relic_tree.delete(row)

        for i, relic in enumerate(filtered):
            tags = ("oddrow",) if i % 2 == 0 else ("evenrow",)
            self.relic_tree.insert(
                "",
                "end",
                values=(
                    relic["name"],
                    relic["id"],
                    relic["total"],
                    relic["wins"],
                    relic["win_rate"],
                    relic["description"],
                ),
                tags=tags,
            )

    def sort_relics(self, key):
        reverse = getattr(self, "_relic_sort_reverse", False)
        self._relic_sort_reverse = not reverse

        def get_val(relic):
            val = relic.get(key)
            if isinstance(val, str):
                return val.lower()
            return val if val else 0

        self.relic_data.sort(key=get_val, reverse=reverse)
        self.filter_relics()

    def on_relic_select(self, event):
        selection = self.relic_tree.selection()
        if selection:
            item = self.relic_tree.item(selection[0])
            values = item["values"]
            if values:
                self.relic_desc_label.config(text=f"Description: {values[5]}")

    def refresh(self):
        self.find_and_load_data()

    def show_help(self):
        help_text = """How to find your STS2 save folder:

WINDOWS:
%LOCALAPPDATA%\\SlayTheSpire2\\steam\\<profile>\\saves\\history

MAC:
~/Library/Application Support/SlayTheSpire2/steam/<profile>/saves/history

LINUX:
~/.local/share/SlayTheSpire2/steam/<profile>/saves/history

NOTE: <profile> is usually a number like 76561198054269638

If auto-detection fails, click "Browse" and select the folder containing your .run files."""
        messagebox.showinfo("Help - Finding STS2 Save Files", help_text)


if __name__ == "__main__":
    app = STSCardViewer()
    app.mainloop()
