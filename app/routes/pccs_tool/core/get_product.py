from config import *
import requests
import json

def get_product(sku):
    """
    Get products from the API.
    """
    try:
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
            verify=False,
            headers={'Content-Type': 'application/json'},
            data=json.dumps(json_data)
        )

        if response.status_code == 200:
            response_data = response.json()
            product_info = response_data.get("products", {}).get(sku, {}).get("productHierarchy", {})
            marketing_category = product_info.get("marketingCategory", {}).get("name")

            if marketing_category == "Workstations":
                marketing_sub_category = product_info.get("marketingSubCategory", {}).get("name")
                
                if marketing_sub_category == "HP HP Mobile Workstation":
                    return marketing_sub_category
                else:
                    return marketing_category
            else:
                return marketing_category
        else:
            print(f"Request failed with status code {response.status_code}")
            print(f"Response: {response.text}")
            return None
    except requests.RequestException as e:
        print(f"Request failed: {e}")
        return None
    except Exception as e:
        print(f"Error: {e}")
        return None
