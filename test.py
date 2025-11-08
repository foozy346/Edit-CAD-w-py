import base64

with open("light-js.ico", "rb") as f:
    encoded = base64.b64encode(f.read()).decode("utf-8")

with open("light-js.txt", "w") as f:
    f.write(encoded)
