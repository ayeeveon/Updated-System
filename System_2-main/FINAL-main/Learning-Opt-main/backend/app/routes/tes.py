import importlib

required = [
    'smtplib',
    'email.mime.text',
    'email.mime.multipart',
    'flask',
    'logging',
    'datetime'
]

missing = []
for lib in required:
    try:
        importlib.import_module(lib)
    except ImportError:
        missing.append(lib)

if missing:
    print(f"❌ Missing libraries: {missing}")
    print("Install with: pip install " + " ".join(missing))
else:
    print("✅ All required libraries are available")