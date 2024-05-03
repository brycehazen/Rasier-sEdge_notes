import os
import subprocess
import sys

def create_virtual_environment(env_name):
    """ Create a virtual environment with the given name. """
    if sys.platform == 'win32':
        subprocess.call(['python', '-m', 'venv', env_name])
    else:
        subprocess.call(['python3', '-m', 'venv', env_name])
    print(f"Virtual environment created at {env_name}")

def install_packages(env_name):
    """ Install required packages using pip from the requirements.txt file. """
    pip_executable = os.path.join(env_name, 'bin', 'pip') if sys.platform != 'win32' else os.path.join(env_name, 'Scripts', 'pip')
    subprocess.call([pip_executable, 'install', '-r', 'requirements.txt'])
    print("Required packages have been installed.")

if __name__ == "__main__":
    env_name = "env"  # Name of the virtual environment
    create_virtual_environment(env_name)
    install_packages(env_name)
