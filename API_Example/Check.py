import resend

resend.api_key = "re_xxxxxx"
response = resend.Emails.get(email_id="8a0ec338-63be-4bb1-bf77-0ff9e190b809")
print(response)


