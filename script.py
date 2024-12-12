import requests  # Import the requests library
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox
from st_aggrid import AgGrid, GridOptionsBuilder

# Dropbox App Credentials
APP_KEY = "w6nrz3ghlfskn5i"  # Replace with your Dropbox App Key
APP_SECRET = "uq94leubvg2xc23"  # Replace with your Dropbox App Secret
REFRESH_TOKEN = "sl.u.AFbLKe_s1qjfmL48cPaEOhqwX5FK-veyXBJJaxejYMPZgqgj7BXq_7kXOV0IqasmG37Ud3oNXZT11hWbzGkp2x1pHcRj1pP5b3LMwpS1YXLWMjbvoBwOUyvH3I7lVHwfvpgcA9t2LG91yEmK_4qc7PXs1nab4lyGLdGTURmpUSkSPmSUYLd5cR-lVxOn6bTVNQ5RAobQoEfdLFZnVCu4f-CBgDSnkS7fFkgg5rVhH0uPU4zv5ojO_MV2ar_UM2NEl8-HnNML-L7omJtuRYtlUE1hZtlT6wEUC7MniHnNPJaizMw5XSs7Xz1iTKvd-y1L1rhBvPSUZHV-u_Na8VkiRDZQo5FPYLfzhsrGckp1P3l2ZezBPVvJ8i9hQddEtF5j1_nnyMjk4FfYfr5gUX1GGnz1sb7JoY-p_1lhpwKPbYP45dPsEbrLt2bPEturl0oIpJB3sQ0Rg7b4FfeWvTE1pMv8oX88hrJJoh_OqO4vS_jj15DtmTHyJNE4XPc0uFmIAUNf5ZU_i6CEnW3MMzSUhSqCRS1UNDQKHAYFHnTvsdKRQURGQ1Lu0g-YoZl5Xpy9RbivObTi0Q6EWoOfoBG8aWtlsJJUyxSLPukgLeFZlTCzLiSiwd7EX01VWJldK6xR07XCYV4qncI9wknSct2i1FT21pzXdkjie5j4qC8T27IXSM39HG-5e5vH26N7cBxVLqARSMcNsBbMxNPiWGSPUuBaTxX-88oUsi4q8IMpebcaCVEBF6Qs-nYBr87K4g5jP961HXv1uocjD8uraq_Kj2xfiJsYm4rk5rTrgp8EdtszhSxUbkdp0kqiru6U68AisgGsXwTCxQCk1a3arJofjXGM5leYx3xOwaWwmLmDnFgxTGo5goj-RUOShgjmWjicxLjNRkhSNcPe4IzyAvghXZL9njDYuGuyIR7_RUQ5clkSTogg2UQ0_MbIGztcbL_339dk_wHVwvqAyuHW0ckwoH28K6A4Lqsrg-Cd-_G_zpbD9ADVp3KZt1yg_DkKri4XKt3SFIikGlwSszyc7V2CQ3JKIC3vaDVhtjkHdfu0xX9gRZL8zPvNM_0CqebcptV8A03pzAVx5wqdYAuYDdRnI2vALZ5N_sKlOAGze7CrjfXy8Nmzl4Ej54U4O8CHadth-_j3rqHBCKvJGM6MpPw86xowNIen4vK9LDDPdCHftmeNiDJBv8uk78maGKEtq99SRoJF5xPx4dGthhpL7R8ofCaZ1Arql5attyydWAG7eToQH4E155UXbPSLXwLSZSu6HppIHngu_cHNqL-cWyJ8Tka1n33pkWENXhw7XbEqDWEBvFXzN1_BFSlm0pyo8cu3yA6Hjv20B3b2skNa5gJIQmiwvGuC8TKKjT1E69cOigDwAWj5D6qW5mv-QcbKHTkUR3-fRyOVDF84QIKH7wNvMiTWgkb2C7RhWJLl8ozdHM--F9lKg_Rk0P8uHOJsHEqWvFk"  # Replace with your Dropbox Refresh Token
PROJECTS_FILE_PATH = "/Project_Data/projects_data_weekly.xlsx"
HR_FILE_PATH = "/Project_Data/Human Resources.xlsx"

# Function to fetch Dropbox access token dynamically
def get_access_token():
    """Fetch a new access token using the refresh token."""
    url = "https://api.dropboxapi.com/oauth2/token"
    data = {
        "grant_type": "refresh_token",
        "refresh_token": REFRESH_TOKEN,
    }
    auth = (APP_KEY, APP_SECRET)
    response = requests.post(url, data=data, auth=auth)  # Fixed issue with missing requests import
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        st.error(f"Error refreshing Dropbox token: {response.json()}")
        raise Exception(f"Token refresh failed: {response.text}")

# Dropbox Functions
def download_from_dropbox(file_path):
    """Download a file from Dropbox."""
    try:
        access_token = get_access_token()
        dbx = dropbox.Dropbox(access_token)
        metadata, res = dbx.files_download(file_path)
        return pd.ExcelFile(res.content)
    except dropbox.exceptions.ApiError as e:
        if e.error.is_path() and e.error.get_path().is_not_found():
            return None
        else:
            st.error(f"Error downloading file: {e}")
            return None

def upload_to_dropbox(df, dropbox_path):
    """Upload a DataFrame to Dropbox as an Excel file."""
    try:
        access_token = get_access_token()
        dbx = dropbox.Dropbox(access_token)
        with pd.ExcelWriter("temp.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        with open("temp.xlsx", "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode("overwrite"))
        st.success(f"Data successfully uploaded to {dropbox_path}.")
    except Exception as e:
        st.error(f"Error uploading data: {e}")
        raise

# Main Application
def main():
    # Your existing code here

    if __name__ == "__main__":
    main()
