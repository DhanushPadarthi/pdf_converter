# PDF to Word OCR Converter - Cloud Deployment

## Quick Deploy to Heroku (Free)

### Prerequisites
- Git installed
- Heroku account (free at heroku.com)
- Heroku CLI installed

### Deploy Steps

1. **Login to Heroku**
```bash
heroku login
```

2. **Create Heroku app**
```bash
heroku create your-pdf-converter-name
```

3. **Set buildpacks**
```bash
heroku buildpacks:add --index 1 heroku/python
heroku buildpacks:add --index 2 https://github.com/heroku/heroku-buildpack-apt
```

4. **Deploy**
```bash
git add .
git commit -m "Deploy PDF converter"
git push heroku main
```

5. **Open your app**
```bash
heroku open
```

Your app will be live at: `https://your-pdf-converter-name.herokuapp.com`

## Deploy to Replit (Super Easy)

1. Go to [replit.com](https://replit.com)
2. Click "Create Repl"
3. Choose "Import from GitHub" or "Upload folder"
4. Upload the `pdf_converter_website` folder
5. Click "Run"
6. Share the generated URL!

## Local Network Access

Change `app.py` to allow network access:
```python
app.run(debug=False, host='0.0.0.0', port=5000)
```

Then share: `http://[YOUR-IP]:5000`

## Environment Variables (For Cloud)

Set these in your cloud platform:
- `FLASK_ENV=production`
- `PYTHONPATH=/app`
