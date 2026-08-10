<div align="center">
<img width="1200" height="475" alt="GHBanner" src="https://github.com/user-attachments/assets/0aa67016-6eaf-458a-adb2-6e31a0763ed6" />
</div>

# Run and deploy your AI Studio app

This contains everything you need to run your app locally.

View your app in AI Studio: https://ai.studio/apps/9d65ffb0-b22c-4775-a76c-f412022e050a

## Run Locally

**Prerequisites:**  Node.js


1. Install dependencies:
   `npm install`
2. Set the `GEMINI_API_KEY` in [.env.local](.env.local) to your Gemini API key
3. Run the app:
   `npm run dev`

## Build and Preview

1. Build production output:
   `npm run build`
2. Preview the built site locally:
   `npm run preview`

## Deploy to GitHub Pages

1. Ensure the repository is set up on GitHub.
2. Run:
   `npm run deploy`
3. Your app will be published at:
   `https://mperdom1.github.io/Genmo-GT-Monitoring-/`
