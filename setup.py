from setuptools import setup, find_packages

setup(
    name="desktopApp",
    version="1.0",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        # Add dependencies here (from requirements.txt)
    ],
    entry_points={
        'console_scripts': [
            'desktopApp = src.desktopApp:main',  # Entry point of your application
        ],
    },
)