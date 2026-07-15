# Email Generator

A powerful Python tool to generate personalized HTML emails in bulk or manually. Perfect for businesses that need to create templated emails for multiple recipients with custom data.

## Features

- **Batch Processing**: Import Excel or CSV files and generate emails for hundreds of contacts automatically
- **Manual Entry**: Create individual emails with custom company names, recipient names, and email addresses
- **Smart Template System**: Use HTML templates to define email layouts with automatic personalization
- **Dual Interface**: Choose between an intuitive GUI or command-line interface
- **Flexible**: Supports both `.xlsx` and `.csv` input files with automatic header detection
- **Safe**: Handles multiple character encodings to prevent encoding errors

## What This Tool Does

This tool takes your contact data and email templates, then combines them to create personalized HTML emails. Each email is saved as a separate `.html` file that you can open in any browser or email client.

**Example**: If you have 50 companies with different templates, this tool can generate 50 personalized emails in seconds instead of manually creating each one.

## Requirements

- Python 3.7 or higher
- Dependencies listed in `requirements.txt`

## Installation

### Step 1: Install Python
Download Python from [python.org](https://www.python.org/downloads/) if you haven't already.

### Step 2: Clone or Download the Project
```bash
git clone https://github.com/AhmadFirdous4/email-generator.git
cd email-generator
```

### Step 3: Install Dependencies
```bash
pip install -r requirements.txt
```

## Quick Start

### Option 1: GUI Mode (Recommended for Beginners)

Simply run the script:
```bash
python email_generator.py
```

A window will open with two tabs:
- **Batch Process**: Upload a data file and process multiple emails at once
- **Manual Entry**: Create a single email with custom details

### Option 2: CLI Mode

For advanced users:
```bash
python email_generator.py --cli
```

This opens an interactive command-line interface for batch processing.

## How to Use

### 1. Set Up Your Templates Folder

Create a folder named `templates` in the same directory as `email_generator.py`:

```
email-generator/
├── email_generator.py
├── templates/
│   ├── enterprise_template.html
│   ├── customer_support.html
│   └── sales_inquiry.html
└── Generated_Emails/
```

### 2. Create HTML Templates

Each `.html` file in the `templates` folder is treated as a template. Templates can include these placeholders:

- `{To_Block}` - Automatically filled with recipient email(s) / Tip: Use in first line to denote the recipient(s)
- `{First_Name}` - Replaced with the recipient's first name / Tip: Use after greeting
- `{Company}` - Replaced with the company name / Tip: Use in subject line as 'Your company - {Company}'

**Example Template** (`templates/enterprise_template.html`):
```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Enterprise Solutions</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            color: #333333;
            line-height: 1.6;
            margin: 0;
            padding: 20px;
            background-color: #f9f9f9;
        }
        .container {
            max-width: 600px;
            margin: 0 auto;
            background-color: #ffffff;
            padding: 30px;
            border-radius: 6px;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        }
        .recipient-block {
            background-color: #f5f5f5;
            padding: 10px 15px;
            margin-bottom: 20px;
            border-left: 3px solid #0066cc;
            border-radius: 3px;
            font-size: 13px;
        }
        .subject-line {
            font-size: 14px;
            font-weight: bold;
            margin-bottom: 20px;
            color: #000000;
            padding-bottom: 10px;
            border-bottom: 2px solid #0066cc;
        }
        .content {
            margin-bottom: 20px;
        }
        .content p {
            margin: 12px 0;
        }
        .content ul {
            margin: 15px 0;
            padding-left: 20px;
        }
        .content li {
            margin: 8px 0;
        }
        .footer {
            font-size: 12px;
            color: #777777;
            margin-top: 30px;
            border-top: 1px solid #eeeeee;
            padding-top: 15px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="recipient-block">
            {To_Block}
        </div>

        <div class="subject-line">
            Subject: Partnership Overview - {Company}
        </div>

        <div class="content">
            <p>Dear {First_Name},</p>

            <p>Thank you for your interest in our enterprise solutions. We are excited to share how our platform can integrate with <strong>{Company}</strong> to streamline your workflow.</p>

            <p><strong>Core Services Included:</strong></p>
            <ul>
                <li>Automated reporting and analytics</li>
                <li>Cloud infrastructure management</li>
                <li>Dedicated technical support</li>
            </ul>

            <p>Please let us know when you have time for a brief introductory call.</p>

            <p>Best regards,<br>
            <strong>[Your Company Name]</strong></p>
        </div>

        <div class="footer">
            <p>This is an automated message generated by the Email Generator application.</p>
        </div>
    </div>
</body>
</html>
```

### Template Best Practices

- Keep it clean and professional
- Test with a preview before generating many emails
- Make sure placeholder names are spelled exactly as shown (case-sensitive)
- One template per email type (e.g., one for sales, one for support)

### 3. Prepare Your Data File

Create an Excel (`.xlsx`) or CSV (`.csv`) file with the following columns:

| Company | Category | First Name | Employee Name | Company Mail | Employee Mail |
|---------|----------|------------|----------------|--------------|----------------|
| Alpha Corp | enterprise_template | John | - | john@alpha.com | - |
| Beta Inc | customer_support | Sarah | - | support@beta.com | - |

**Column Notes:**
- `Company` - The company name (required)
- `Category` - Must match your template file names (required)
- `First Name` - Used for personalization (optional)
- `Employee Name` - Alternative to First Name (optional)
- `Company Mail` & `Employee Mail` - Email addresses (optional, multiple emails supported)

### 4. Generate Emails

**For Batch Processing:**
1. Open the GUI: `python email_generator.py`
2. Click "Browse" and select your data file
3. Click "Process File"
4. Check the `Generated_Emails` folder for your HTML files

**For Manual Entry:**
1. Open the GUI and click the "Manual Entry" tab
2. Fill in the company name, contact name, and emails
3. Select a template category
4. Click "Generate Email"

## Output

All generated emails are saved as `.html` files in the `Generated_Emails` folder with names like:
- `Alpha Corp_sales_inquiry.html`
- `Beta Inc_customer_support.html`

You can:
- Open them in your browser to preview
- Send them via email (copy-paste the HTML or use "Send as HTML")
- Edit them before sending
- Use them in email marketing platforms

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "No templates found" | Create a `templates` folder in the same directory as the script |
| File encoding errors | The tool supports UTF-8, Windows-1252, and Latin-1 encodings automatically |
| Category not matching | Ensure your CSV/Excel category column matches the template filename (without `.html`) |
| Missing emails in output | Check that your data file has Company Mail or Employee Mail columns |

## File Structure

```
email-generator/
├── email_generator.py          # Main application
├── requirements.txt             # Python dependencies
├── README.md                    # This file
├── .gitignore                   # Files to exclude from Git
├── templates/                   # Your HTML templates go here
│   └── your_template.html
└── Generated_Emails/            # Output folder (created automatically)
    └── Company_Category.html
```

## Contributing

Found a bug or have a suggestion? Feel free to open an issue!

## Support

For questions or issues:
1. Check the Troubleshooting section above
2. Review the template examples
3. Open an issue on GitHub

---

**Happy email generating!**
