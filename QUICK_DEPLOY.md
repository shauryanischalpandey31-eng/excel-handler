# Quick Deployment Guide

## ⚠️ Important: Netlify doesn't support Django

Netlify is for static sites only. Django needs a Python server. Use **Render** instead (free tier available).

## Deploy to Render (5 minutes)

### Step 1: Push to GitHub
```bash
git add .
git commit -m "Prepare for deployment"
git push origin main
```

### Step 2: Deploy on Render

1. **Go to Render**: https://render.com
2. **Sign up** with GitHub
3. **Click "New +"** → **"Web Service"**
4. **Connect your GitHub repository**
5. **Configure**:
   - **Name**: `excel-handler` (or any name)
   - **Environment**: `Python 3`
   - **Build Command**: 
     ```bash
     pip install -r requirements.txt && python manage.py collectstatic --noinput && python manage.py migrate
     ```
   - **Start Command**: 
     ```bash
     gunicorn blast_project.wsgi:application
     ```
6. **Add Environment Variables**:
   - `SECRET_KEY`: Generate one using:
     ```python
     python -c "from django.core.management.utils import get_random_secret_key; print(get_random_secret_key())"
     ```
   - `DEBUG`: `False`
   - `ALLOWED_HOSTS`: `your-app-name.onrender.com` (Render will show you the URL)
7. **Click "Create Web Service"**
8. **Wait 5-10 minutes** for deployment

### Step 3: Your App is Live!

Your app will be available at: `https://your-app-name.onrender.com/excel/`

---

## Alternative: Railway (Even Easier)

1. Go to https://railway.app
2. Sign up with GitHub
3. Click "New Project" → "Deploy from GitHub repo"
4. Select your repository
5. Railway auto-detects Django and deploys!
6. Add environment variables in "Variables" tab
7. Done! Get your URL

---

## Need Help?

Check `DEPLOYMENT.md` for detailed instructions and troubleshooting.

