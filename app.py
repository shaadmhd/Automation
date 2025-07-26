from flask import Flask, request, jsonify
from automate_leads import initialize_driver, login_to_portal, add_new_lead

app = Flask(__name__)

@app.route('/submit-lead', methods=['POST'])
def submit_lead():
    lead_data = request.json
    if not lead_data:
        return jsonify({"status": "error", "message": "No lead data received"}), 400

    driver = initialize_driver()
    if not driver:
        return jsonify({"status": "failed", "message": "WebDriver failed"}), 500

    if not login_to_portal(driver, "9600879549", "vasan@naruto"):
        driver.quit()
        return jsonify({"status": "failed", "message": "Login failed"}), 401

    success = add_new_lead(driver, lead_data)
    driver.quit()

    if success:
        return jsonify({"status": "success", "message": "Lead submitted"})
    else:
        return jsonify({"status": "failed", "message": "Lead submission failed"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
