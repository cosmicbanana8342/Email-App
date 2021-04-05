# -*- coding: utf-8 -*-
"""
Created on Fri Jan 17 08:08:48 2020

@author: sherangagamwasam
"""
import configure

from mailchimp3 import MailChimp
from string import Template
from tkinter import messagebox

# =============================================================================
# authentication
# =============================================================================

MC_API_KEY = configure.MC_API_KEY
MC_USER_NAME = configure.MC_USER_NAME

client = MailChimp(
        mc_api = MC_API_KEY,
        mc_user = MC_USER_NAME)


# =============================================================================
# add members to the existing audience function
# =============================================================================
def add_members_to_audience_function(
        audience_id, 
        email_list, 
        client=client):
        
        audience_id = audience_id
        email_list = email_list

        if len(email_list)!=0:
            flag = 0
            for email_iteration in email_list:
                try:
                    data = {
                        "status": "subscribed",
                        "email_address": email_iteration
                    }

                    client.lists.members.create(list_id=audience_id, data=data)

                except Exception as error:
                    messagebox.showerror("Something went wrong!", error)
                    flag = 1
            if flag != 1:
                messagebox.showinfo("Contact(s) Added.", "Contact(s) added successfully.")

        else:
            messagebox.showinfo("No email to add.", "Email list is empty.")


# =============================================================================
# campaign creation function
# =============================================================================
def campaign_creation_function(campaign_name, audience_id, from_name, reply_to, preview_text, client=client):
        
        try:
            campaign_name = campaign_name
            audience_id = audience_id
            from_name = from_name
            reply_to = reply_to

            data = {
                "recipients" :
                {
                    "list_id": audience_id
                },
                "settings":
                {
                    "subject_line": campaign_name,
                    "from_name": from_name,
                    "reply_to": reply_to,
                    "preview_text": preview_text
                },
                "type": "regular"
            }

            new_campaign = client.campaigns.create(data=data)
            
            return new_campaign
        except Exception as error:
            messagebox.showerror("Something went wrong!", error)
            return "ok"
    
# =============================================================================
# template creation
# =============================================================================

def customized_template(html_code, campaign_id, client = client):
        
        html_code = html_code
        campaign_id = campaign_id

        string_template = Template(html_code).safe_substitute()
        
        try:
            client.campaigns.content.update(
                    campaign_id=campaign_id,
                    data={'message': 'Campaign message', 'html': string_template}
                    )
        except Exception as error:
            messagebox.showerror("Failed to add template!", error)
            return "ok"

# =============================================================================
# send the mail campaign
# =============================================================================

def send_mail(campaign_id, client = client):
        
        try:
            client.campaigns.actions.send(
                    campaign_id = campaign_id
                )
        except Exception as error:
            messagebox.showerror("Failed to send mail", error)
            return "ok"
