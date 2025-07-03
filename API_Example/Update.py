import resend
from datetime import datetime, timedelta

resend.api_key = "re_xxxxxx"

minute_from_now = (datetime.now() + timedelta(minutes=10)).isoformat()

update_params: resend.Emails.UpdateParams = {
  "id": "8a0ec338-63be-4bb1-bf77-0ff9e190b809",
  "scheduled_at": minute_from_now
}

resend.Emails.update(params=update_params)