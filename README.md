# download & install python
# git clone or download repository https://github.com/tuanzw/rwsbn.git to folder %userprofile%

# Install
cd %userprofile%\rwsbn
python -m venv .
Scripts\activiate
pip install -r requirements.txt

# change email in .env file
# correct anywhere password (base64 encode); url_site; site_name; doc_library; sharepoint_folder
# open cmd & run command: mkdir [attachment_folder]; mkdir [attachment_folder]\[attachement_move_to_folder]; mkdir [attachement_folder]\[sp_download_folder]
# update sender list for session: # sender list per trucking vendor by ; seperate
# create desktop shortcut of rwsbn.bat
# change .ico file (icon file of shortcut) 

# Mailbox
# create folder under Inbox\[email_folder]
# create folder under Inbox\[email_folder]\[email_move_to_folder]