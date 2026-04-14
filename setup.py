from cx_Freeze import setup, Executable
import platform
versions = '1.2.7'

build_exe_options = {
    "packages": ["comtypes"],
    "include_files": ['resource/'],
    "build_exe": f"EggManager_{versions}"
}

executable_kwargs = {
    "script": 'EggManager_GUI.py',
    "target_name": f'EggManager_{versions}',
    "icon": 'resource/eggui.ico',
}

if platform.system() == "Windows":
    executable_kwargs["base"] = "gui"
    executable_kwargs["uac_admin"] = True

exe = [Executable(**executable_kwargs)]
 
setup(
    name='EggManager',
    version = versions,
    author='TUVup',
    options = {"build_exe": build_exe_options},
    executables = exe
)
