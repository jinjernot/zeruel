from app.config.config import *
import requests
import json

def get_product(sku):
    """Get products from API"""
    try:
        ca_cert = ca_cert_path
        client_cert = (client_cert_path, client_key_path)

        json_data = {
            "sku": [sku],
            "countryCode": "US",
            "languageCode": "EN",
            "layoutName": "PDPCOMBO",
            "requestor": "HERMESQA-PRO",
            "reqContent": ["hierarchy"]
        }

        response = requests.post(
            url,
            cert=client_cert,
            verify=ca_cert,
            headers={'Content-Type': 'application/json'},
            data=json.dumps(json_data)
        )

        if response.status_code == 200:
            # Save the response to a JSON file
            #json_filename = f"response_{sku}.json"
            #with open(json_filename, 'w') as json_file:
            #    json.dump(response.json(), json_file)

            response_data = response.json()
        
            marketing_category = response_data["products"][sku]["productHierarchy"]["marketingCategory"].get("name")

            if marketing_category == "Workstations":
                marketing_sub_category = response_data["products"][sku]["productHierarchy"]["marketingSubCategory"].get("name")
                
                if marketing_sub_category == "HP Mobile Workstation":
                    # Use marketing_sub_category value instead of marketing_category
                    api_data = marketing_sub_category
                else:
                    api_data = marketing_category
            else:
                api_data = marketing_category
            return api_data

        else:
            print(f"Request failed with status code {response.status_code}")
            print(f"Response: {response.text}")
    except Exception as e:
        print(f"Error: {e}")
