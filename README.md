# pgmVideoClips

Welcome to **pgmVideoClips**, a Python program that processes video clips using a simple guide. This README will help you get started with the installation, setup, and usage of the program.

## Installation

1. **Install or Update pip:** Make sure you have the latest version of `pip` installed.
 ```shell
 py -m pip install --upgrade pip
 py -m pip --version
 ```
  2. **Installing a Virtual Environment:** To keep your project dependencies isolated, let's set up a virtual environment using virtualenv
  ```python
  py -m pip install --user virtualenv
  ```
3. **Create and Activate Virtual Environment:** Navigate to your project's directory and create a virtual environment. If you're using Python 2, replace venv with virtualenv in the commands.
  ```python
  py -m venv env
  ```
  Activate your virtual environment:
  -  On Windows (Command Prompt)
  ```python
  .\env\Scripts\activate
  ```
  - On Windows (PowerShell)
  ```python
  .\env\Scripts\Activate.ps1
  ```
4. **Install Dependencies:** Install the required dependencies listed in the 'requirements.txt' file.
```python
py -m pip install -r requirements.txt
```
# Usage

1. **Preparing Files:** Place your spreadsheet and '.mp4' files into the '/Main' folder.
2. **Running the Program:** Execute the program according to your needs.

# Demo
![pgmVIdeoClipperdemo](https://user-images.githubusercontent.com/59584473/169143437-36be22ba-f48d-4922-8ab2-cb2a5ccd6af2.gif)

# Spreadsheet Example
![pgmVideoClipper_capture](https://user-images.githubusercontent.com/59584473/169146465-1a331c3e-1100-4bd7-b131-7ded0f9e1640.png)
*Captured during encoding progress*

## Troubleshooting

If you encounter any issues during the installation or usage of the program, feel free to [open an issue](link-to-issues-page) on our GitHub repository.

## Contributing

We welcome contributions to enhance `pgmVideoClips`. If you'd like to contribute, please follow our [contribution guidelines](link-to-contribution-guidelines).

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT)
