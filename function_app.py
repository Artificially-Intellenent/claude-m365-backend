import azure.functions as func
import requests
import json
from datetime import datetime, timedelta
import logging
import os

app = func.FunctionApp()

# Configuration - SET THESE WITH YOUR CREDENTIALS
TENANT_ID = os.environ.get("TENANT_ID")  # Your Tenant ID
CLIENT_ID = os.environ.get("CLIENT_ID")  # Your Application (Client) ID
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")  # Your Client Secret
FIELD_GUY_EMAIL = os.environ.get("FIELD_GUY_EMAIL", "")  # Field guy's email address

# Get access token for Microsoft Graph
def get_access_token():
    """Get OAuth token for Microsoft Graph API"""
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    body = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    
    response = requests.post(url, headers=headers, data=body)
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        logging.error(f"Token error: {response.text}")
        return None

# Get calendar events
def get_calendar_events(email, days_ahead=30):
    """Fetch calendar events for the next N days"""
    token = get_access_token()
    if not token:
        return {"error": "Failed to get access token"}
    
    headers = {"Authorization": f"Bearer {token}"}
    
    # Date range for query
    start_time = datetime.utcnow().isoformat() + "Z"
    end_time = (datetime.utcnow() + timedelta(days=days_ahead)).isoformat() + "Z"
    
    # Microsoft Graph API call
    url = f"https://graph.microsoft.com/v1.0/users/{email}/calendarview"
    params = {
        "startDateTime": start_time,
        "endDateTime": end_time,
        "$orderby": "start/dateTime",
        "$top": 100
    }
    
    response = requests.get(url, headers=headers, params=params)
    
    if response.status_code == 200:
        events = response.json().get("value", [])
        # Format for Excel/Claude
        formatted = []
        for event in events:
            formatted.append({
                "subject": event.get("subject", ""),
                "start": event.get("start", {}).get("dateTime", ""),
                "end": event.get("end", {}).get("dateTime", ""),
                "location": event.get("location", {}).get("displayName", ""),
                "attendees": len(event.get("attendees", [])),
                "isReminderOn": event.get("isReminderOn", False)
            })
        return {"events": formatted, "count": len(formatted)}
    else:
        logging.error(f"Calendar error: {response.text}")
        return {"error": f"Failed to fetch calendar: {response.status_code}"}

# Get recent emails
def get_recent_emails(email, num_emails=20):
    """Fetch recent emails"""
    token = get_access_token()
    if not token:
        return {"error": "Failed to get access token"}
    
    headers = {"Authorization": f"Bearer {token}"}
    
    # Microsoft Graph API call
    url = f"https://graph.microsoft.com/v1.0/users/{email}/messages"
    params = {
        "$orderby": "receivedDateTime desc",
        "$top": num_emails,
        "$select": "subject,from,receivedDateTime,bodyPreview,isRead,importance"
    }
    
    response = requests.get(url, headers=headers, params=params)
    
    if response.status_code == 200:
        emails = response.json().get("value", [])
        # Format for Excel/Claude
        formatted = []
        for email_msg in emails:
            formatted.append({
                "subject": email_msg.get("subject", ""),
                "from": email_msg.get("from", {}).get("emailAddress", {}).get("name", ""),
                "received": email_msg.get("receivedDateTime", ""),
                "preview": email_msg.get("bodyPreview", "")[:100],  # First 100 chars
                "isRead": email_msg.get("isRead", False),
                "importance": email_msg.get("importance", "normal")
            })
        return {"emails": formatted, "count": len(formatted)}
    else:
        logging.error(f"Email error: {response.text}")
        return {"error": f"Failed to fetch emails: {response.status_code}"}

# Create calendar event
def create_calendar_event(email, subject, start_time, end_time, location=""):
    """Create a new calendar event"""
    token = get_access_token()
    if not token:
        return {"error": "Failed to get access token"}
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    body = {
        "subject": subject,
        "start": {
            "dateTime": start_time,
            "timeZone": "UTC"
        },
        "end": {
            "dateTime": end_time,
            "timeZone": "UTC"
        },
        "location": {
            "displayName": location
        }
    }
    
    url = f"https://graph.microsoft.com/v1.0/users/{email}/events"
    response = requests.post(url, headers=headers, json=body)
    
    if response.status_code == 201:
        return {"success": True, "eventId": response.json().get("id")}
    else:
        logging.error(f"Create event error: {response.text}")
        return {"error": f"Failed to create event: {response.status_code}"}

# Main HTTP-triggered function
@app.route(route="m365_data", methods=["GET", "POST"])
def m365_backend(req: func.HttpRequest) -> func.HttpResponse:
    """Main backend function - handles requests from Excel/Claude"""
    
    try:
        # Get request parameters
        req_body = req.get_json() if req.method == "POST" else {}
        action = req.args.get("action") or req_body.get("action")
        email = req.args.get("email") or req_body.get("email") or FIELD_GUY_EMAIL
        
        if not email:
            return func.HttpResponse(
                json.dumps({"error": "Email address required"}),
                status_code=400,
                mimetype="application/json"
            )
        
        # Route to appropriate function
        if action == "calendar":
            days = int(req_body.get("days", 30))
            result = get_calendar_events(email, days)
        
        elif action == "emails":
            num = int(req_body.get("num", 20))
            result = get_recent_emails(email, num)
        
        elif action == "create_event":
            subject = req_body.get("subject")
            start = req_body.get("start")
            end = req_body.get("end")
            location = req_body.get("location", "")
            
            if not all([subject, start, end]):
                return func.HttpResponse(
                    json.dumps({"error": "subject, start, and end are required"}),
                    status_code=400,
                    mimetype="application/json"
                )
            
            result = create_calendar_event(email, subject, start, end, location)
        
        else:
            return func.HttpResponse(
                json.dumps({"error": "Unknown action. Use: calendar, emails, or create_event"}),
                status_code=400,
                mimetype="application/json"
            )
        
        return func.HttpResponse(
            json.dumps(result),
            status_code=200,
            mimetype="application/json"
        )
    
    except Exception as e:
        logging.error(f"Error: {str(e)}")
        return func.HttpResponse(
            json.dumps({"error": str(e)}),
            status_code=500,
            mimetype="application/json"
        )
