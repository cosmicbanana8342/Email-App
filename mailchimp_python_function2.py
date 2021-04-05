import mailchimp_marketing as MailchimpMarketing
from mailchimp_marketing.api_client import ApiClientError
import json
import configure
import hashlib
from tkinter import messagebox

client = MailchimpMarketing.Client()

# authentication
MC_API_KEY = configure.MC_API_KEY
MC_SERVER = configure.MC_SERVER
client.set_config({
  "api_key": MC_API_KEY,
  "server": MC_SERVER
})

# function to delete a member
def del_member(list_id,email):
	subscriber_hash = hashlib.md5(email.encode('utf-8')).hexdigest()
	try:
		client.lists.delete_list_member_permanent(list_id, subscriber_hash)
	except ApiClientError as error:
		response = messagebox.showerror("Something went wrong!", error.text)
		return(response)

def unsubscribe(list_id,email):
	email = email
	member_email_hash = hashlib.md5(email.encode('utf-8')).hexdigest()
	try:
		unsubscribe_response = client.lists.update_list_member(list_id, member_email_hash, {"status": "unsubscribed"})
	except ApiClientError as error:
		unsubscribe_response = messagebox.showerror("Something went wrong!", error.text)
	return unsubscribe_response

def subscribe(list_id,email):
	email = email
	member_email_hash = hashlib.md5(email.encode('utf-8')).hexdigest()
	try:
		subscribe_response = client.lists.update_list_member(list_id, member_email_hash, {"status": "subscribed"})
	except ApiClientError as error:
		subscribe_response = messagebox.showerror("Something went wrong!", error.text)
		return subscribe_response


# function to get all members info in json
def members_info(list_id):
	try:
		email_list = {}
		members_info_json = client.lists.get_list_members_info(list_id, count=1000)
		# json_object = json.dumps(members_info_json, indent = 4)
		# with open("members_info_response.json", "w") as outfile:
		# 	outfile.write(json_object)
		members = members_info_json['members']
		for member in members:
			email_list[member['email_address']] = member['status']
		return email_list

	except ApiClientError as error:
		response = messagebox.showerror("Something went wrong!", error.text)
		return(response)