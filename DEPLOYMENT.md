# Deployment Guide for Excel Handler Django App

## Deploy to Render (Recommended - Free Tier Available)

### Step 1: Prepare Your Repository
1. Make sure all your code is committed to a Git repository (GitHub, GitLab, or Bitbucket)
2. Ensure you have a `requirements.txt` file with all dependencies
3. Make sure `render.yaml` and `Procfile` are in your repository root

### Step 2: Deploy to Render

1. **Sign up/Login to Render**
   - Go to https://render.com
   - Sign up with GitHub/GitLab/Bitbucket

2. **Create New Web Service**
   - Click "New +" → "Web Service"
   - Connect your repository
   - Select the repository with your Django app

3. **Configure the Service**
   - **Name**: `excel-handler-django` (or any name you prefer)
   - **Environment**: `Python 3`
   - **Build Command**: 
     ```bash
     pip install -r requirements.txt && python manage.py collectstatic --noinput && python manage.py migrate
     ```
   - **Start Command**: 
     ```bash
     gunicorn blast_project.wsgi:application
     ```

4. **Set Environment Variables**
   - `SECRET_KEY`: Generate a new secret key (you can use: `python -c "from django.core.management.utils import get_random_secret_key; print(get_random_secret_key())"`)
   - `DEBUG`: `False`
   - `ALLOWED_HOSTS`: Your Render URL (e.g., `excel-handler-django.onrender.com`)
   - `PYTHON_VERSION`: `3.12.6`

5. **Deploy**
   - Click "Create Web Service"
   - Render will build and deploy your app
   - Wait for deployment to complete (5-10 minutes)

6. **Your App URL**
   - Once deployed, you'll get a URL like: `https://excel-handler-django.onrender.com`
   - Access your app at: `https://excel-handler-django.onrender.com/excel/`

### Step 3: Set Up Database (Optional - for production)

If you want to use PostgreSQL instead of SQLite:
1. In Render dashboard, go to "New +" → "PostgreSQL"
2. Create a new PostgreSQL database
3. Copy the "Internal Database URL"
4. Add environment variable: `DATABASE_URL` with the copied URL
5. Your app will automatically use PostgreSQL

### Step 4: Set Up Media Files Storage

For production, consider using:
- **AWS S3** for media file storage
- **Cloudinary** for image/media hosting
- Or keep using local storage (files will be lost on redeploy)

## Alternative: Deploy to Railway

1. Go to https://railway.app
2. Sign up with GitHub
3. Click "New Project" → "Deploy from GitHub repo"
4. Select your repository
5. Railway will auto-detect Django and deploy
6. Add environment variables in the "Variables" tab
7. Your app will be live at a Railway URL

## Alternative: Deploy to Heroku

1. Install Heroku CLI
2. Run:
   ```bash
   heroku create your-app-name
   heroku config:set SECRET_KEY=your-secret-key
   heroku config:set DEBUG=False
   heroku config:set ALLOWED_HOSTS=your-app-name.herokuapp.com
   git push heroku main
   ```

## Important Notes

- **Static Files**: Handled by WhiteNoise middleware
- **Media Files**: Consider using cloud storage for production
- **Database**: SQLite works for small apps, PostgreSQL recommended for production
- **Secret Key**: Never commit your secret key to Git
- **Debug Mode**: Always set to `False` in production

## Troubleshooting

- If static files don't load: Run `python manage.py collectstatic` locally and commit the `staticfiles` folder
- If database errors: Make sure migrations are run (`python manage.py migrate`)
- If 500 errors: Check Render logs for detailed error messages

