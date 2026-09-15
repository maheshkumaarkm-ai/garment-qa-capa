# Garment QA CAPA – Android

This is the first native Android version of the Garment QA CAPA app.

Features:
- Select/take a garment defect photo.
- Factory and date fields.
- OpenAI image analysis.
- Editable defect, root cause and corrective action.
- PowerPoint generation using the supplied QA/CAPA layout.
- Defect photo fitted into the left panel.
- PowerPoint text uses wrapping/autofit settings.
- GitHub Actions builds an APK automatically.

IMPORTANT:
- For this personal prototype, the user enters their own OpenAI API key in the app.
- Do not hard-code an API key into source code or commit it to GitHub.
- A production/company version should move OpenAI calls to a secure backend.

BUILD:
1. Upload this project to GitHub.
2. Open Actions -> Build Android APK -> Run workflow.
3. Download the generated artifact and install the APK on Android.
