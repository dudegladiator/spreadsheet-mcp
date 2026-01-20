#!/usr/bin/env python3
"""Check Google Cloud setup and API access."""

import json
from pathlib import Path

def main():
    print("\n=== Google Sheets MCP - Setup Check ===\n")
    
    # 1. Check credentials file
    cred_paths = [
        Path("./credentials/service-account.json"),
        Path("./credentials/service-cred.json"),
        Path("./service-account.json"),
        Path("./service-cred.json"),
    ]
    
    cred_file = None
    for p in cred_paths:
        if p.exists():
            cred_file = p
            break
    
    if not cred_file:
        print("❌ No credentials file found!")
        print("   Place your service-account.json in ./credentials/")
        return
    
    print(f"✓ Credentials file found: {cred_file}")
    
    # 2. Read service account email
    with open(cred_file) as f:
        creds = json.load(f)
    
    email = creds.get("client_email", "Unknown")
    project = creds.get("project_id", "Unknown")
    
    print(f"✓ Service Account: {email}")
    print(f"✓ Project ID: {project}")
    
    # 3. Test API access
    print("\n--- Testing API Access ---\n")
    
    try:
        from src.spreadsheet_mcp.auth import get_sheets_service
        service = get_sheets_service()
        print("✓ Google Sheets API client initialized")
    except Exception as e:
        print(f"❌ Failed to initialize API client: {e}")
        return
    
    # 4. Try to create a spreadsheet (tests both Sheets and Drive API)
    print("\n--- Testing Create Spreadsheet ---\n")
    
    try:
        result = service.spreadsheets().create(
            body={"properties": {"title": "MCP Test - Delete Me"}}
        ).execute()
        
        spreadsheet_id = result["spreadsheetId"]
        url = result["spreadsheetUrl"]
        
        print(f"✓ Successfully created test spreadsheet!")
        print(f"  URL: {url}")
        print(f"\n⚠️  Don't forget to delete this test spreadsheet!")
        
    except Exception as e:
        error_str = str(e)
        print(f"❌ Failed to create spreadsheet: {e}\n")
        
        if "not have permission" in error_str or "403" in error_str:
            print("This usually means one of the following:")
            print(f"\n1. Google Sheets API is not enabled for project '{project}'")
            print(f"   → Go to: https://console.cloud.google.com/apis/library/sheets.googleapis.com?project={project}")
            print("   → Click 'Enable'")
            print(f"\n2. Google Drive API is not enabled (needed to create new spreadsheets)")
            print(f"   → Go to: https://console.cloud.google.com/apis/library/drive.googleapis.com?project={project}")
            print("   → Click 'Enable'")
            print("\n3. After enabling, wait 1-2 minutes for changes to propagate")
        
        return
    
    print("\n✅ All checks passed! Your MCP server is ready to use.")


if __name__ == "__main__":
    main()
