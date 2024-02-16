import os
import subprocess
import platform


def sys_check():
    # Получаем информацию о текущей операционной системе
    os_system = platform.system()
    return os_system


def env_check():
    activate_file = "activate"

    if 'Windows' in sys_check():

        venv_path = "venv"
        if os.path.isdir(venv_path) and os.path.isfile(os.path.join(venv_path + '\\Scripts\\',
                                                                    activate_file)):
            print("Виртуальное окружение уже установлено.")
        else:
            print("Виртуальное окружение не установлено, создаём...")
            create_venv_command = "python -m venv venv"
            subprocess.run(create_venv_command, shell=True)

    elif 'Linux' in sys_check():

        venv_path = "venv"
        if os.path.isdir(venv_path) and os.path.isfile(os.path.join(venv_path + '/Scripts/',
                                                                    activate_file)):
            print("Виртуальное окружение уже установлено.")
        else:
            print("Виртуальное окружение не установлено, создаём...")
            create_venv_command = "python -m venv venv"
            subprocess.run(create_venv_command, shell=True)

    else:
        print('Для отличных от Windows и Linux операционных систем программа не реализована')


def moduls_install():
    virtualenv_python_path = '.\\venv\\Scripts\\python.exe'
    library_to_install = "openpyxl"
    install_command = [virtualenv_python_path, "-m", "pip", "install", library_to_install]
    subprocess.run(install_command)
