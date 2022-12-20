from requests import request
# step 1: fill in the client credentials
CLIENT_ID = "608569353527-ffsa66hes94f1ergsnmltpgiqrfbscef.apps.googleusercontent.com"
CLIENT_SECRET = "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"
# step 2: get a temporary access code manually
# https://accounts.google.com/o/oauth2/auth?access_type=offline&approval_prompt=auto&client_id=608569353527-ffsa66hes94f1ergsnmltpgiqrfbscef.apps.googleusercontent.com&response_type=code&scope=https://www.googleapis.com/auth/spreadsheets&redirect_uri=http://localhost
ACCESS_CODE = "4/0AdQt8qie_Yn5TERZxFzhNF90ZY5jJY82BIvkXzEXjJ-P_gFFAcl95ropkNAxGs9laGl3Xg"
# step 3: run this script
def main():
    # get refresh token/access token from access code
    url = "https://accounts.google.com/o/oauth2/token"
    data = {
        "grant_type": "authorization_code",
        "code": ACCESS_CODE,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://www.googleapis.com/auth/spreadsheets",
        "redirect_uri": "http://localhost"
    }
    headers = {
        "content-type": "application/x-www-form-urlencoded"
    }
    r = request("POST", url, data=data, headers=headers)
    print(r.text)
if __name__ == "__main__":
    main()