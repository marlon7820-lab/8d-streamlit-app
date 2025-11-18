import re

path = "app.py"

with open(path, "r", encoding="utf-8") as f:
    text = f.read()

# Replace ALL non-breaking spaces with normal spaces
cleaned = text.replace("\u00A0", " ")

with open(path, "w", encoding="utf-8") as f:
    f.write(cleaned)

print("All U+00A0 characters removed!")
