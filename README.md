# BodyCount

BodyCount is a Fusion 360 add-in created for and owned by Vermland, an interior design firm specializing in handcrafted modular design. It serves as a bridge between Fusion 360 and the Kitchen X Excel spreadsheet, automatically populating project data to enable accurate pricing calculations.

## Installation

1. **Download and extract the add-in**: Download the BodyCount folder and extract it to the location where you want it installed.

2. **Open Fusion 360** and go to **Utilities** > **Add-Ins** > **Scripts and Add-Ins**.

3. In the **Scripts and Add-Ins** dialog, click the dropdown arrow next to the **+** button in the toolbar.

4. Select **"Script or add-in from device"** to browse for the add-in folder.

5. Navigate to and select the extracted BodyCount folder, then click **Open**.

6. **Restart Fusion 360**. The add-in will now appear in the Add-Ins list.

7. Click **Run** to enable the add-in for the first time. Optionally, check **Run on Startup** to auto-load it when Fusion 360 starts.

## First-Time Setup

When you run BodyCount for the first time, it will automatically install required Python packages (pyserde, pypiwin32, and exceltypes). A terminal window may briefly appear during this process. After installation completes, you will see a message asking you to restart Fusion 360. Please restart the application to complete the setup.
