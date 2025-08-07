# Google Slides Automation Instructions

## Option 1: Google Apps Script (Recommended)

### Step-by-Step Setup:

1. **Open Google Apps Script**
   - Go to [script.google.com](https://script.google.com)
   - Click "New Project"

2. **Replace Default Code**
   - Delete the existing `myFunction()` code
   - Copy and paste the entire contents of `google-apps-script.js` from this repository

3. **Verify Presentation ID**
   - The script is already configured with your presentation ID: `1xmY2iPUIQm0dbrGDWl-rLQDEr_kMDuH8rAvfL08mQyI`
   - If you need to use a different presentation, update the `PRESENTATION_ID` variable

4. **Save and Name Your Project**
   - Click the save icon or Ctrl+S
   - Give your project a name like "Management Slides Automation"

5. **Grant Permissions**
   - Click "Run" button next to `populateSlides`
   - You'll be prompted to authorize the script
   - Click "Review Permissions" → "Allow"

6. **Run the Script**
   - Click "Run" again
   - Check the execution log for success message
   - Your slides should now be populated!

### Troubleshooting:

- **"Script can't access presentation"**: Make sure you have edit access to the Google Slides presentation
- **"Presentation not found"**: Double-check the presentation ID in the URL
- **"Permission denied"**: Ensure you've authorized the script to access Google Slides

## Option 2: Manual Copy-Paste Guide

If you prefer not to use the script, here's an optimized manual process:

### Quick Copy-Paste Workflow:

1. **Open both files:**
   - Your Google Slides presentation
   - `Management-Slides-Content.md` from this repository

2. **For each slide section in the markdown:**
   - Create a new slide in Google Slides
   - Copy the slide title (the `**bold text**`)
   - Paste as the slide title
   - Copy the bullet points
   - Paste into the slide body
   - Copy any quotes (*italic text*)
   - Add as a text box at the bottom

3. **Suggested slide layouts:**
   - Title slide: "Title and Subtitle"
   - Content slides: "Title and Body"
   - Quote-heavy slides: "Title and Two Columns"

### Time Estimate:
- Script method: 2-3 minutes (one-time setup)
- Manual method: 15-20 minutes

## Option 3: PowerPoint Import

If you have PowerPoint available:

1. I can create a `.pptx` file with the content
2. You can then import it directly into Google Slides
3. Google Slides → File → Import slides → Upload PowerPoint file

Let me know if you'd like me to create the PowerPoint version!

---

**Recommendation:** Try the Google Apps Script method first - it's the most automated and will save significant time, especially if you need to make updates later.