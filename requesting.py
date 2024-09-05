import requests

url = 'https://trujillo.bluence.com/rest'
api_key = 'adsouifapo8udjrviummhmanpidg7niyprrttrq2'

# Set up headers
headers = {
    'Authorization': f'{api_key}',  # or 'API-Key': api_key depending on the API
}

# Make the GET request with headers
response = requests.get(url, headers=headers, timeout = 10000)

# Check response
if response.status_code == 200:
    data = response.json()  # Assuming the response is in JSON format
    print(data)
else:
    print(f"Error: {response.status_code}")
