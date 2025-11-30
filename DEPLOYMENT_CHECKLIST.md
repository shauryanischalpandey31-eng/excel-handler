# Deployment Checklist for Excel Handler App

## ✅ Step 1: Code Pushed to GitHub
- [x] All changes committed
- [x] Pushed to `main` branch
- [x] Repository: `shauryanischalpandey31-eng/excel-handler`

## Step 2: Render Deployment

### Option A: Auto-Deploy (If Already Connected)
1. Go to https://dashboard.render.com
2. Find your service: `excel-handler-tiyc` (or similar)
3. Check "Events" tab - should show "Deploying..." or "Live"
4. Wait 5-10 minutes for deployment to complete

### Option B: Manual Deploy (If Not Connected)
1. Go to https://render.com
2. Sign in with GitHub
3. Click "New +" → "Web Service"
4. Connect repository: `shauryanischalpandey31-eng/excel-handler`
5. Configure:
   - **Name**: `excel-handler`
   - **Environment**: `Python 3`
   - **Build Command**: 
     ```bash
     pip install -r requirements.txt && python manage.py collectstatic --noinput && python manage.py migrate
     ```
   - **Start Command**: 
     ```bash
     gunicorn blast_project.wsgi:application
     ```
6. Add Environment Variables:
   - `SECRET_KEY`: Generate using:
     ```python
     python -c "from django.core.management.utils import get_random_secret_key; print(get_random_secret_key())"
     ```
   - `DEBUG`: `False`
   - `ALLOWED_HOSTS`: `excel-handler-tiyc.onrender.com` (or your Render URL)
   - `PYTHON_VERSION`: `3.12.6` (optional)
7. Click "Create Web Service"

## Step 3: Verify Environment Variables

In Render dashboard, go to your service → "Environment" tab, verify:

- ✅ `SECRET_KEY` is set (not the default insecure key)
- ✅ `DEBUG` is `False`
- ✅ `ALLOWED_HOSTS` includes your Render URL
- ✅ No missing variables

## Step 4: Check Deployment Status

1. Go to "Events" or "Logs" tab in Render
2. Look for:
   - ✅ "Build successful"
   - ✅ "Deploy successful"
   - ✅ "Service is live"
3. If errors:
   - Check build logs for missing dependencies
   - Check runtime logs for application errors
   - Verify all environment variables are set

## Step 5: Test Deployed App

1. Visit your app URL: `https://excel-handler-tiyc.onrender.com`
2. Should redirect to: `https://excel-handler-tiyc.onrender.com/excel/`
3. Test upload:
   - Upload an Excel file
   - Click "Process"
   - Complete workflow
   - Verify charts display correctly
4. Check charts:
   - Charts show exact Excel values
   - Tooltips show correct numbers
   - Table displays historical vs predicted

## Step 6: Verify Chart Data Fix

1. Upload Excel file with monthly data
2. Process through workflow
3. Check browser Network tab:
   - Look for `process_all_workflows` response
   - Verify `chart_data` field exists
   - Check `chart_data.products[0].historical` has correct values
4. Verify charts:
   - Hover over points - tooltips show exact values
   - Table shows correct historical/predicted split
   - No placeholder values

## Troubleshooting

### Issue: Build Fails
- **Check**: `requirements.txt` has all dependencies
- **Fix**: Add missing packages to requirements.txt
- **Check**: Python version compatibility

### Issue: App Crashes on Startup
- **Check**: Environment variables are set
- **Check**: `ALLOWED_HOSTS` includes Render URL
- **Check**: Database migrations completed
- **Check**: Static files collected

### Issue: Charts Not Showing
- **Check**: Browser console for JavaScript errors
- **Check**: Network tab for chart_data in response
- **Check**: Chart.js library loaded correctly

### Issue: 500 Errors
- **Check**: Render logs for detailed error
- **Check**: `DEBUG=False` but errors not showing
- **Temporary**: Set `DEBUG=True` to see errors (then change back)

## Post-Deployment

- [ ] Test file upload
- [ ] Test chart rendering
- [ ] Test download functionality
- [ ] Verify tooltips show correct values
- [ ] Check table displays correctly
- [ ] Test with different Excel file formats

## Your App URL

Once deployed, your app will be available at:
**https://excel-handler-tiyc.onrender.com/excel/**

## Need Help?

- Check Render logs: Dashboard → Your Service → "Logs" tab
- Check Django logs: Look for extraction debug messages
- Browser DevTools: Network tab and Console tab

