# 🚀 Deploy BA Task Manager to Railway.app

Railway.app offers **free tier hosting** perfect for this application. Deploy in **5 minutes**.

## Step 1: Create a Railway Account

1. Go to [railway.app](https://railway.app)
2. Click **Sign Up** (recommend GitHub login)
3. Connect your GitHub account

## Step 2: Upload Your Code to GitHub

If you don't have a GitHub account, create one at [github.com](https://github.com).

### Option A: Using GitHub Web Interface (Easiest)

1. Go to [github.com/new](https://github.com/new)
2. Name it `ba-task-manager`
3. Make it **Public** (so Railway can access it)
4. Click **Create Repository**
5. Upload files:
   - Click **Add file → Upload files**
   - Select all files from your folder:
     - `app.py`
     - `index.html`
     - `requirements.txt`
     - `Procfile`
     - `runtime.txt`
     - `.gitignore`
   - Commit the files

### Option B: Using Git Command Line

```bash
# Clone the repo you created
git clone https://github.com/YOUR_USERNAME/ba-task-manager.git
cd ba-task-manager

# Copy all your files here
# Then:
git add .
git commit -m "Initial commit: BA Task Manager"
git push origin main
```

## Step 3: Deploy on Railway

1. Go to [railway.app/dashboard](https://railway.app/dashboard)
2. Click **New Project → Deploy from GitHub**
3. Select your `ba-task-manager` repository
4. Wait for automatic deployment (2-3 minutes)
5. Your app gets a **live URL** like `ba-task-manager.up.railway.app`

## Step 4: Access Your Live App

Visit your Railway deployment URL in browser. You now have a **live, publicly accessible app!**

## Setting Environment Variables (Optional)

If you want to customize the data directory:

1. Go to your Railway project
2. Click **Environment Variables**
3. Add:
   ```
   DATA_DIR=./data
   FLASK_ENV=production
   PORT=8000
   ```
4. Redeploy

## Free Tier Limits

- **Storage**: 5GB
- **Bandwidth**: Unlimited
- **CPU**: Limited but sufficient for this use case
- **Uptime**: 99.9%
- **No credit card required**

## Troubleshooting

**App crashes after deploy?**
- Check logs in Railway dashboard
- Common issue: Excel file not initialized on first run
- Solution: App auto-creates the Excel file, just wait and refresh

**Can't upload Excel file?**
- Railway provides persistent storage at `./data` folder
- Excel file is created automatically on first task/project creation

**Data lost after redeploy?**
- Railway keeps data in persistent volume, it won't be lost
- Your Excel file stays in `./data/ba_task_manager.xlsx`

## Updating Your App

Made changes locally? Update Railway:

```bash
git add .
git commit -m "Update: [describe changes]"
git push origin main
```

Railway auto-redeploys on every push!

## Custom Domain (Optional)

Want `mytaskmanager.com` instead of `ba-task-manager.up.railway.app`?

1. Buy a domain on Namecheap or GoDaddy
2. In Railway project → **Settings → Domain**
3. Add your custom domain
4. Update DNS at your domain provider

## Monitoring & Logs

In Railway dashboard:
- **Deployments** tab: See deployment history
- **Logs** tab: Real-time app logs
- **Metrics** tab: CPU, Memory, Network usage

## Scale Your App (Paid)

Once you grow:
- **Pay as you go**: Only $5/month for 1000+ users
- **Auto-scaling**: Railway scales automatically
- **Database upgrade**: Add PostgreSQL or MongoDB if needed (currently using Excel, can migrate later)

---

**Your app is now live and publicly accessible!** 🎉

Share your URL: `https://ba-task-manager-XXXXX.up.railway.app`

Anyone with the URL can access it. To make it private, add authentication (contact if needed).

