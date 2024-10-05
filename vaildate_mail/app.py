from flask import Flask, render_template, request, jsonify
import re
import smtplib
import dns.resolver

app = Flask(__name__)

# Validate email syntax using regex
def is_valid_email_syntax(email):
    regex = r'^[^\s@]+@[^\s@]+\.[^\s@]+$'
    return re.match(regex, email) is not None

# Get MX record for the domain
def get_mx_record(domain):
    try:
        mx_records = dns.resolver.resolve(domain, 'MX')
        return mx_records[0].exchange.to_text()  # Return first MX record
    except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN, dns.resolver.NoNameservers):
        return None

# Perform SMTP validation to check if email exists
def validate_email_smtp(email):
    global server
    domain = email.split('@')[1]

    # Get MX record for the domain
    mx_record = get_mx_record(domain)
    if not mx_record:
        return False, "No MX record found for domain"

    # Set up SMTP connection
    try:
        server = smtplib.SMTP(timeout=10)
        server.set_debuglevel(0)

        # SMTP handshake
        server.connect(mx_record)
        server.helo(server.local_hostname)
        server.mail('test@example.com')
        code, message = server.rcpt(email)

        if code == 250:
            return True, "Email address is valid"
        else:
            return False, f"Email address is not valid"

    except smtplib.SMTPServerDisconnected:
        return False, "SMTP server disconnected unexpectedly"
    except smtplib.SMTPConnectError:
        return False, "Failed to connect to SMTP server"
    except smtplib.SMTPRecipientsRefused:
        return False, "Recipient refused"
    except Exception as e:
        return False, f"Unexpected error: {str(e)}"
    finally:
        if 'server' in locals():
            server.quit()

# Main email validation function
def validate_email(email):
    if not is_valid_email_syntax(email):
        return False, "Invalid email format"

    return validate_email_smtp(email)

# Route for form input and URL query handling
@app.route('/', methods=['GET', 'POST'])
def index():
    email = None
    result = None

    # Check if the form is submitted
    if request.method == 'POST':
        email = request.form.get('email')
    # Check if email is provided via URL
    elif request.args.get('email'):
        email = request.args.get('email')

    if email:
        is_valid, message = validate_email(email)
        result = {'email': email, 'is_valid': is_valid, 'message': message}

    return render_template('index.html', result=result)

# Route for validating email via API (JSON response)
@app.route('/api/validate', methods=['GET'])
def api_validate_email():
    email = request.args.get('email')
    if not email:
        return jsonify({'error': 'No email provided'}), 400

    is_valid, message = validate_email(email)
    return jsonify({'email': email, 'is_valid': is_valid, 'message': message})

if __name__ == '__main__':
    app.run(debug=False)
