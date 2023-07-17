# ADSKDashboard

Shortcuts to open Autodesk Software.

If you have multiple Autodeskproducts in diffrent releases and languages installade it can be a pain to open exactly what you need. At least thats what I was experiencing and so I built this tool.

Currently it works for AutoCAD / AutoCAD Mechanical / Inventor / Inventor Read Only from Release 2017 to 2021.
# Update v0.96

- updated to scan for 2024 products

# Update v0.95

- updated to scan for 2023 products

# Update v0.94

- added detection for languagepacks inventor 2022 updates

# Update v0.93

- updated to scan for 2022 products

# Update v0.92

- Fixed a bug that did show ReadOnly for Version that do not have ReadOnly
- Auto Close Function (After opening something, it closes down automatically)
- Closes down if the window looses focus

# Tool does not start!

[Have a look here](https://github.com/TWiesendanger/ADSKDashboardPS#tool-doesnt-start)

# Table of Contents

- [ADSKDashboard](#adskdashboard)
- [Update v0.96](#update-v096)
- [Update v0.95](#update-v095)
- [Update v0.94](#update-v094)
- [Update v0.93](#update-v093)
- [Update v0.92](#update-v092)
- [Tool does not start!](#tool-does-not-start)
- [Table of Contents](#table-of-contents)
- [What is installed?](#what-is-installed)
- [Settings](#settings)
- [Closing](#closing)
- [Tool doesn't start](#tool-doesnt-start)
- [License](#license)

# What is installed?

At startup it will check what is installed and depending on that it will display a Icon or not.

![](/docs/adskd_interface1.png)

So as a sample if click on 2020 there are alot more products displayed because there is also AutoCAD and AutoCAD Mechanical installed.

![](/docs/adskd_interface2.png)

# Settings

At the moment there are only three settings. you can change to a dark mode and you can set the window to be always on top.
Also you can decide if the window should be minimized after clicking on something to start.

![](/docs/adskd_settings.png)

# Closing

If you hover over the "X" in the top right corner, you can see that it tells you close to systray. It wont actualy close from there.

![](/docs/adskd_closetosystray.png)

After that you can restart it from systray.

![](/docs/adskd_systrayicon.png)

# Tool doesn't start

At the moment it is not signed so you probably get a windows defender message which you need to allow.

![](/docs/adskd_smartscreen.png)

Also make sure that all dll files are not blocked. Open res\assembly folder and check by rightclicking on dll file.

![](/docs/adskd_blocked.jpg)

# License

MIT License

Copyright (c) 2020 Tobias Wiesendanger

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
