from cx_Freeze import setup, Executable

# ADD FILES
add_files = [
    ("NssmUI.exe", "")
]

# TARGET
target = Executable(
    script="HwName.py",
    base="Win32GUI",
    icon="HwName.ico",
    uac_admin=True,
    target_name="HwName.exe"
)

# SETUP CX FREEZE
setup(
    name="Hardware Name Modifier",
    version="0.1.2025.0530",
    description="Hardware Name Modifier for Windows",
    author="Pikachu Ren",
    options={
        'build_exe': {
            "include_msvcr": True,
            'include_files': add_files,
            "packages": [
                "ttkbootstrap.utility",
                "ttkbootstrap",
            ],
        },
    },
    executables=[target],
)
