import json

with open('.local/share/opencode/tool-output/tool_d65a784ed001DPkG6D9SZFGelZ') as f:
    relics = json.load(f)

RELIC_INFO = {}
for r in relics:
    rid = r['id']
    desc = r['description'].replace('[blue]', '').replace('[/blue]', '').replace('[gold]', '').replace('[/gold]', '').replace('\n', ' ')
    RELIC_INFO[rid] = {
        'name': r['name'],
        'description': desc[:200],
        'rarity': r.get('rarity', 'Unknown')
    }

print("RELIC_INFO = {")
for k, v in sorted(RELIC_INFO.items()):
    esc_name = v['name'].replace('"', '\\"')
    esc_desc = v['description'].replace('"', '\\"').replace('\n', ' ')
    print(f'    "{k}": {{"name": "{esc_name}", "description": "{esc_desc}", "rarity": "{v["rarity"]}"}},')
print("}")
