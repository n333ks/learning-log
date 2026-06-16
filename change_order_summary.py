import json

# Load the JSON file
with open("change_order_procedures.json") as f:
    data = json.load(f)

# Print document header
print("=== CHANGE ORDER PROCEDURES SUMMARY ===")
print(f"Company: {data['document']['company']}")
print(f"Last Reviewed: {data['document']['last_reviewed']}")
print("")

# Loop through each part and its scenarios
for part in data['parts']:
    print(f"--- {part['title']} ---")
    for scenario in part['scenarios']:
        print(f"  [{scenario['id']}] {scenario['title']}")
        print(f"       Status: {scenario['status']}")
        print(f"       Friction: {scenario['friction_level']}")
    print("")

print("=== NON-NEGOTIABLES ===")
for rule in data['non_negotiables']:
    print(f"  - {rule['rule']}")

print("")
print("=== HIGH FRICTION SCENARIOS ===")
for part in data['parts']:
    for scenario in part['scenarios']:
        if scenario['friction_level'] == "high":
            print(f"  [{scenario['id']}] {scenario['title']}")
            print(f"       {scenario['status']}")