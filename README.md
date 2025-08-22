# OpenAI Usage Tracker

A Google Apps Script that automatically tracks OpenAI API usage and costs across all your organization's projects.

## 🚀 Features

- **Daily Usage Tracking**: Automatically fetches yesterday's usage data
- **Cost Monitoring**: Tracks costs for yesterday + today
- **Project Management**: Creates individual sheets for each project
- **Duplicate Prevention**: Combines duplicate entries and prevents data duplication
- **Automatic Updates**: Can run daily at midnight automatically
- **Cost Calculation**: Shows actual costs and costs with 30% markup

## 📋 Prerequisites

- Google Sheets
- OpenAI Organization Admin API Key
- Google Apps Script access

## ⚙️ Setup

### 1. Create Google Sheet
- Create a new Google Sheet
- Open **Extensions → Apps Script**

### 2. Add Script
- Copy the `script.gs` code into your Apps Script editor
- Save the project

### 3. Set API Key
- In Apps Script, go to **Project Settings**
- Add a new script property:
  - **Property**: `OPENAI_ADMIN_KEY`
  - **Value**: Your OpenAI Organization Admin API Key

### 4. Run Script
- Go back to your Google Sheet
- Refresh the page
- You'll see a new menu: **OpenAI Usage**
- Click **OpenAI Usage → Run now**

## 📊 What Gets Created

### Summary Sheets
- **Summary Usage**: Daily usage data by project and model
- **Summary Costs**: Daily cost data by project and line item

### Project Sheets
- **Project Name - Usage**: Individual usage data per project
- **Project Name - Costs**: Individual cost data per project

## 🎯 Menu Options

- **Run now**: Fetch latest data (usage: yesterday, costs: 2 days)
- **Get yesterday data**: Get only yesterday's data
- **Install daily trigger**: Set up automatic daily execution at midnight
- **Test API connection**: Verify your API key works
- **Get Organization Info**: View your OpenAI organization details
- **Show Tracked Projects**: List all projects in your organization
- **Clean Duplicate Data**: Remove any existing duplicate entries

## ⏰ Daily Updates

### Manual
- Run **OpenAI Usage → Run now** daily

### Automatic
- Run **OpenAI Usage → Install daily trigger** once
- Script runs automatically every day at midnight

## 📈 Data Structure

### Usage Data Columns
1. Date
2. Project ID
3. Project Name
4. Model
5. Total Tokens (Input+Output)
6. Input tokens
7. Output tokens
8. Cached tokens
9. Requests
10. Cost (USD)

### Cost Data Columns
1. Date
2. Project ID
3. Project Name
4. Line Item
5. Amount (USD) - Hidden
6. Cost (+30%) - Visible

## 🔧 Customization

### Track Specific Projects
Edit the `TRACK_SPECIFIC_PROJECTS` array in the script:
```javascript
const TRACK_SPECIFIC_PROJECTS = [
  'proj_your_project_id_here',
  'proj_another_project_id'
];
```

### Time Zones
The script automatically uses your Google Sheet's timezone setting.

## 🚨 Troubleshooting

### API Connection Issues
- Verify your `OPENAI_ADMIN_KEY` is correct
- Use **Test API connection** to check
- Ensure your API key has organization access

### No Data Showing
- Check the execution logs in Apps Script
- Verify your projects have recent activity
- Run **Run now** to fetch fresh data

### Duplicate Data
- Use **Clean Duplicate Data** to remove duplicates
- The script now prevents future duplicates automatically

## 📝 Notes

- **Usage data**: Only yesterday (not today)
- **Cost data**: Yesterday + today (2 days)
- **Duplicate handling**: Automatically combines duplicate entries
- **Cost calculation**: Shows both actual cost and cost with 30% markup
- **Project sheets**: Created automatically for each active project

## 🤝 Support

If you encounter issues:
1. Check the execution logs in Apps Script
2. Verify your API key and permissions
3. Ensure your OpenAI organization has active projects

## 📄 License

This project is open source and available under the MIT License.
