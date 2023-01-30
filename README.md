# da-tool

da-tool is a script that allows you to gather data and perform various administrative operations on Knet profiles and AM Console profiles inside of Amazon's network. These features include:

- Searching Knet via employee ID or username
- Searching AM Console via transporter ID
- Saving the name, profile links, email address, DSP, and onboarding status
- Downloading badge photos
- Resetting Knet passwords

## Important

In order for this script to work, you will need to:

1. Be connected to the company VPN network.
2. Download _geckodriver.exe_ and place it into the assets folder.
3. Edit _da.py_ to add DSP names and 4-letter codes to the DSP_MAP object.

## Disclaimer

I do not take responsibility for any damages that may be incurred through the use of this tool. Audit the code yourself and use at your own risk.
