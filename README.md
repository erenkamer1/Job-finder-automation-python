# Automated Job Application Assistant

This project is a tool that automates email collection and sending processes for job seekers.

## Features

- 🔍 Automatic Company Email Finding
  - Scanning company websites using Google search results
  - Smart email pattern matching
  - Saving results to Excel

- 📧 Automatic Email Sending
  - Gmail integration
  - Daily email sending limit control
  - Automatic CV attachment
  - Multiple Gmail account support
  - Progress tracking

## Installation

1. Install required Python packages:
```bash
pip install -r requirements.txt
```

2. Chrome browser and ChromeDriver must be installed.

3. Set up your Gmail account:
   - Enable "Less secure app access" in Gmail
   - If using two-factor authentication, create an app password

## Usage

### Email Scraper

```bash
python email_scraper.py
```

- Select a text file containing company list
- Results will be saved to `email_results.xlsx`

### Email Sender

```bash
python email_sender_selenium.py
```

- Enter your Gmail credentials
- Specify your CV file
- Prepare your email template
- Sending will begin

## Configuration

### email_results.xlsx format:
- Company: Company name
- Email: Found email address
- Status: Email finding status
- Target Country: Country where the company is located

### email_tracking.json format:
```json
{
    "email@example.com": {
        "2024-02-09": 5
    }
}
```

## Security Notes

- Never write your Gmail passwords in the code
- Use two-factor authentication
- Use app passwords
- Be careful not to exceed daily email limits

## Contributing

1. Fork this repository
2. Create a new branch (`git checkout -b feature/amazing`)
3. Commit your changes (`git commit -m 'Added amazing feature'`)
4. Push your branch (`git push origin feature/amazing`)
5. Create a Pull Request

## License

This project is licensed under the MIT License. See the `LICENSE` file for details. 