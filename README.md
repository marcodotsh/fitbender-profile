# Repo info

This repository automates extraction of the official [Garmin FIT SDK profile.xlsx](https://github.com/garmin/fit-sdk-tools/blob/main/Profile.xlsx) to JSON format, with [community driven additions](https://www.harryonline.net/blog-en/beyond-the-sdk-uncovering-undocumented-garmin-fit-file-information/).
The github action triggers weekly, checks for changes and updates the profile.json file.

# Credits
Thanks to Harry Oosterveen for maintaining the list of undocumented values in [this Google Sheet](https://docs.google.com/spreadsheets/d/1x34eRAZ45nbi3U3GyANotgmoQfj0fR49wBxmL-oLogc/edit?usp=sharing). Check out its [FIT File Viewer](https://www.fitfileviewer.com/), where you can also contact him in case you discover new fields and you want to share them.