# Powerpoint Slide Timeline

This VBA script generates a timeline indicator/ progress bar on the botton of each slide of your presentation

![Screenshot](/Screenshot.png)

## How to use  :wrench:
---

Step 1: Open PowerPoint

Step 2: Enable Developer tab by going to File → Options → Customise Ribbon. (Under `Customize the Ribbon` select `Developer`and click `OK`)

Step 3: Open `Developer` tab in PowerPoint and click on `Macros` → Type macro name → Click `Create` → Copy & paste everything in `Timeline.bas` file in to the code editor.

Step 4: Chage variable colors according to your PowerPoint theme. (_Tip: Match border color with slide background color_)

```
'------Theme colors------'
'Adjust these to match your power point theme
past = RGB(165, 255, 250)
present = RGB(0, 255, 205)
future = RGB(2, 69, 173)
borders = RGB(7, 32, 69)
```

![Colors guide](/ColorsGuide.png)

Step 5: Finally click `Run Sub` button or press <kbd>F5</kbd>