
target_file = r'c:\Users\cloud\Desktop\EPA-grading\grading-system\app.py'
with open(target_file, 'r', encoding='utf-8') as f:
    lines = f.readlines()
    for i, line in enumerate(lines):
        if 'month' in line.lower() or ' 7 ' in line:
            print(f"{i+1}: {line.strip()}")
