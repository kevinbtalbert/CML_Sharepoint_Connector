import os
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext

# Define SharePoint credentials and site URL
sharepoint_url = 'https://your-org.sharepoint.com/'
username = 'your-username@org.com'
password = 'your-password'
document_library_name = 'Documents'  # Change to your document library name
local_save_path = os.getcwd() + '/data/'

# Authenticate and create a client context
auth_ctx = AuthenticationContext(sharepoint_url)
if auth_ctx.acquire_token_for_user(username, password):
    ctx = ClientContext(sharepoint_url, auth_ctx)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    print(f"Authenticated to SharePoint site: {web.properties['Title']}")

    # Get the document library
    doc_library = ctx.web.lists.get_by_title(document_library_name)
    ctx.load(doc_library)
    ctx.execute_query()

    # Get all items in the document library
    items = doc_library.items
    ctx.load(items)
    ctx.execute_query()

    # Create local save directory if it does not exist
    if not os.path.exists(local_save_path):
        os.makedirs(local_save_path)

    # Iterate through items and download files
    for item in items:
        file_url = item.properties["FileRef"]
        file_name = item.properties["FileLeafRef"]
        file_path = os.path.join(local_save_path, file_name)

        with open(file_path, "wb") as local_file:
            file = ctx.web.get_file_by_server_relative_url(file_url)
            file.download(local_file)
            ctx.execute_query()
            print(f"Downloaded: {file_name} to {file_path}")

else:
    print("Authentication failed")
