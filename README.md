# Office 2021 LTSC Deployment Script

This PowerShell script automates the process of deploying Microsoft Office 2021 LTSC (Long-Term Servicing Channel) following the guidelines provided by [Microsoft's official deployment documentation](https://learn.microsoft.com/en-us/office/ltsc/2021/deploy). It allows you to choose between Office 2021 Standard, Office 2021 ProPlus, or a custom configuration, simplifying the entire deployment process.

## Features

- Downloads the latest Office Deployment Tool directly from Microsoft.
- Supports three configuration options:
  1. **Office 2021 Standard**
  2. **Office 2021 ProPlus**
  3. **Custom Configuration** (requires `Configuration.xml` to be placed in the project directory)
- Automatically handles the extraction and setup process.
- Ensures a unique installation directory for each run to avoid conflicts.

## Prerequisites

- Windows PowerShell 5.0 or higher
- Internet access to download the Office Deployment Tool and configuration files

## How to Use

### 1. **Clone or Download the Repository:**

Clone this repository or download the script to your local machine.

  ```sh
   git clone https://github.com/jdepew88/download_office_ltsc.git
  ```
### 2. **Run the Script:**

Open PowerShell as Administrator and navigate to the directory containing the script.
  ```sh
  cd path\to\download_office_ltsc
  ```
Execute the script:
  ```sh
  .\deploy_office.ps1
  ```

## Choosing Your Configuration.mxl

### 1. **Choose Your Configuration:**
When prompted, choose one of the following options:

  ##### 1:  Office 2021 Standard  -  Installs Office 2021 Standard using the predefined config_standard.xml.
  ##### 2:  Office 2021 ProPlus  -  Installs Office 2021 ProPlus using the predefined config_ProPlus.xml.
  ##### 3:  Custom Configuration  -  Use your own Configuration.xml. Ensure this file is placed in the                 project directory   before running the script.

  ```sh
  Select the Office 2021 version to install:
  1. Office 2021 Standard
  2. Office 2021 ProPlus
  3. Custom (Place your Configuration.xml file in the project directory before running this script)
  Enter the number of your choice (1, 2, or 3):
  ```
 - **Note for Custom Configuration:** If you select option 3 and the **'Configuration.xml'** file is not found, the script will remind you to place it in the project directory and then exit.  (See next step to use Microsoft's OCT 'Office Customization Tool'.  Here you can choose the apps you prefer.

### 2. **Using the Office Customization Tool:**

 - If you prefer to create your own custom configuration file, you can use the [Office Customization Tool](https://config.office.com/deploymentsettings)
.
 - This web-based tool allows you to customize various settings for your Office deployment, including product selection, update channels, languages, and more.
 - After customizing your settings, download the Configuration.xml file and place it in the project directory before running this script.

## What the Script Does
### 1. **Automatic Download and Installation:**

The script will automatically:

 - Download the Office Deployment Tool from Microsoft.
 - Extract the tool into a uniquely named directory.
 - Download the Office 2021 installation files based on your selected configuration.
 - Run the installation using the specified configuration file.

### 2. **Monitor Progress:**

The script will display progress information as it downloads and installs Office 2021.

 - If the installation is successful, a confirmation message will appear.
  ```sh
  Office installation completed successfully.
  ```
## Troubleshooting
 - **Script Execution Policy:**

  If you encounter issues running the script due to execution policy restrictions, you may need to change the policy       temporarily:
  ```sh
  Set-ExecutionPolicy RemoteSigned -Scope Process
  ```

 - **Internet Connectivity:**
    Ensure that you have a stable internet connection, as the script requires downloading files from Microsoft.

 - **Custom Configuration File:**
  If using a custom configuration, double-check that the Configuration.xml file is correctly placed in the project directory before starting the script.

## License
This project is licensed under the MIT License. See the [LICENSE](https://github.com/jdepew88/download_office_ltsc/blob/main/LICENSE) file for more details.

