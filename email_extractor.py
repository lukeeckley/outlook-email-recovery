import argparse
import os
import re
import sys

try:
    import pypff
except ImportError:
    print("[+] Install the libpff Python bindings to use this script")
    sys.exit(1)

message_list = []
messages = 0

def main(pst_file, output_dir):
    print("[+] Accessing {} PST file".format(pst_file))
    pst = pypff.open(pst_file)
    root = pst.get_root_folder()
    print("[+] Traversing PST folder structure")
    recursePST(root)
    print("[+] Identified {} messages..".format(messages))

def recursePST(base):
    for folder in base.sub_folders:
        if folder.number_of_sub_folders:
            recursePST(folder)
        processMessages(folder)

def emailExtractor(item):
	if "<" in item:
		start = item.find("<") + 1
		stop = item.find(">")
		email = item[start:stop]
	else:
		email = item.split(":")[1].strip().replace('"', "")
	if "@" not in email:
		domain = False
	else:
		domain = email.split("@")[1].replace('"', "")
	return email, domain

def processMessages(folder):
    global messages
    print("[+] Processing {} Folder with {} messages".format(folder.name, folder.number_of_sub_messages))
    if folder.number_of_sub_messages == 0:
        return
    for message in folder.sub_messages:
        eml_from, replyto, returnpath = ("", "", "")
        messages += 1
        try:
            headers = message.get_transport_headers().splitlines()
        except AttributeError:
			# No email header
            continue
        for header in headers:
            if header.strip().lower().startswith("from:"):
                eml_from = header.strip().lower()
            elif header.strip().lower().startswith("reply-to:"):
                replyto = header.strip().lower()
            elif header.strip().lower().startswith("return-path:"):
                returnpath = header.strip().lower()
        if eml_from == "" or (replyto == "" and returnpath == ""):
            # No FROM value or no Reply-To / Return-Path value
            continue
        dumpMessage(folder, message, eml_from, replyto, returnpath)

def dumpMessage(folder, msg, eml_from, reply, return_path):
    reply_bool = False
    return_bool = False
    from_email, from_domain = emailExtractor(eml_from)

    if reply != "":
        reply_bool = True
        reply_email, reply_domain = emailExtractor(reply)
    if return_path != "":
        return_bool = True
        return_email, return_domain = emailExtractor(return_path)

    try:
        if msg.html_body is None:
            if msg.plain_text_body is None:
                if msg.rtf_body is None:
                    print("BODY:")
                else:
                    print("BODY: {}".format(msg.rtf_body))
            else:
                print("BODY: {}".format(msg.plain_text_body))
        else:
            print("BODY: {}".format(msg.html_body))
    except:
        print("Coudn't process body")

if __name__ == '__main__':
	# Command-line Argument Parser
	parser = argparse.ArgumentParser(description="Outlook Email Dumper")
	parser.add_argument("PST_FILE", help="File path to input PST file")
	parser.add_argument("OUTPUT_DIR", help="Output Dir")
	args = parser.parse_args()

	if not os.path.exists(args.OUTPUT_DIR):
		os.makedirs(args.OUTPUT_DIR)

	if os.path.exists(args.PST_FILE) and os.path.isfile(args.PST_FILE):
		main(args.PST_FILE, args.OUTPUT_DIR)
	else:
		print("[-] Input PST {} does not exist or is not a file".format(args.PST_FILE))
		sys.exit(4)
