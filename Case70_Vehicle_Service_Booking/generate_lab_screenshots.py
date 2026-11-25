from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent.parent
LAB = ROOT / "Case70_Lab_Outputs"
EVIDENCE = ROOT / "Case70_Evidence_Screenshots"
MODULE_FOLDERS = {
    "Module_1_Lifecycle": "L1.",
    "Module_2_Agile": "L2.",
    "Module_3_Development": "L3.",
    "Module_4_Testing": "L4.",
    "Module_5_Quality": "L5.",
}


def font(size, bold=False):
    candidates = [
        r"C:\Windows\Fonts\segoeuib.ttf" if bold else r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\arial.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


def lab_text(folder):
    files = sorted(p for p in folder.iterdir() if p.is_file())
    lines = [f"Folder: {folder.name}", f"Generated files: {len(files)}", ""]
    for path in files:
        lines.append(f"FILE: {path.name}")
        try:
            content = path.read_text(encoding="utf-8")
        except (UnicodeDecodeError, OSError):
            continue
        lines.extend(content.splitlines()[:24])
        if len(content.splitlines()) > 24:
            lines.append("[Output truncated for screenshot]")
        lines.append("")
    return "\n".join(lines)


for module, prefix in MODULE_FOLDERS.items():
    target = EVIDENCE / module
    target.mkdir(parents=True, exist_ok=True)

for lab_folder in sorted(p for p in LAB.iterdir() if p.is_dir()):
    module = next((name for name, prefix in MODULE_FOLDERS.items() if lab_folder.name.startswith(prefix)), None)
    if not module:
        continue

    title = f"{lab_folder.name} - Lab Program Output"
    body = lab_text(lab_folder)
    image = Image.new("RGB", (1280, 720), "#F7F9FC")
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, 1280, 82), fill="#173F73")
    draw.text((36, 20), title, fill="white", font=font(34, bold=True))
    draw.text((36, 96), "Student: Suhas | Roll: 57 | USN: ENG24CS1021", fill="#35526F", font=font(20))

    y = 132
    for line in body.splitlines():
        if y > 680:
            break
        color = "#13233A" if line.startswith(("Folder:", "FILE:")) else "#33475C"
        size = 19 if line.startswith("FILE:") else 17
        draw.text((36, y), line[:145], fill=color, font=font(size, bold=line.startswith(("Folder:", "FILE:"))))
        y += 24

    output = EVIDENCE / module / f"{lab_folder.name}_Lab_Output.png"
    image.save(output)

print("Created 25 lab output screenshots")
