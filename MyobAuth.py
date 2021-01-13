__author__ = 'Jade McKenzie', 'Tyler McCamley'

import requests
import requests_oauthlib
import configparser
import json
import os
import sys

DEBUG = False
ERROR_PREFIX = '>> Error from MyobAuth >> '
MSG_PREFIX = 'Message from MyobAuth >> '


# Program start is only used for debug purposes
def program_start(run):
    app = EagleMyobApi()

    while run is True:

        entry = input()
        if entry == 'get auth':
            app.get_user_authorisation()
        elif entry == 'check auth':
            app.print_auth()
        elif entry == 'get token':
            app.get_token()
        elif entry == 'check token':
            app.check_token()
        elif entry == 'update token':
            app.update_token()
        elif entry == 'request':
            app.myob_request('companyfile')
        elif entry == 'request overdue':
            app.myob_request('overdue')
        elif entry == 'quit':
            run = False
        else:
            print(MSG_PREFIX + "Not a valid command")
            continue


class EagleMyobApi:

    dev_key = ''
    dev_secret = ''
    dev_redirect_uri = 'http://desktop'
    scope = ''

    retry_counter = 0

    auth_code = None
    auth_state = None

    token_url = 'https://secure.myob.com/oauth2/v1/authorize/'

    config = configparser.ConfigParser()
    curpath = os.path.dirname(os.path.realpath(sys.argv[0]))
    cfgpath = os.path.join(curpath, 'settings.ini')
    config.read(cfgpath)

    access_token = config.get('TOKENS', 'Access')
    refresh_token = config.get('TOKENS', 'Refresh')

    oauth = requests_oauthlib.OAuth2Session(dev_key,
                                            redirect_uri=dev_redirect_uri,
                                            scope=scope)
    auth_url, state = oauth.authorization_url(
        'https://secure.myob.com/oauth2/account/authorize/?')

    def get_user_authorisation(self):

        # TODO: Reduce timeout when user force closes selenium browser.

        try:
            from selenium import webdriver
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as ec
            from webdriver_manager.chrome import ChromeDriverManager

            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_experimental_option("excludeSwitches",
                                                   ['enable-automation'])
            driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)

            driver.get(self.auth_url)
            WebDriverWait(driver, 500).until(ec.title_contains('code'))
            print(MSG_PREFIX + 'Found ' + str(driver.current_url))
            auth_code = driver.title
            driver.quit()

            # Split Auth Code/State
            auth_code = auth_code.split('=')
            self.auth_state = auth_code[2]
            self.auth_code = auth_code[1].strip('state').strip()
            print(MSG_PREFIX + 'Authorisation Retrieved. Checking States.')
            self.check_state()

            # Get initial token
            self.get_token()
        except Exception as e:
            print(ERROR_PREFIX + str(e))

    def check_state(self):
        # This is a security measure to prevent third-party access.
        try:
            if self.auth_state == self.state:
                print(MSG_PREFIX + "States Match")
            else:
                input(MSG_PREFIX + "State Mismatch, force closing program. "
                      "Press any key to close.")
                quit()
        except Exception as e:
            print(ERROR_PREFIX + str(e))

    def print_auth(self):
        # For testing purposes.
        try:
            print(MSG_PREFIX + 'Auth Code: ' + str(self.auth_code))
            print(MSG_PREFIX + 'Auth State: ' + str(self.auth_state))
            pass
        except Exception as e:
            print(ERROR_PREFIX + str(e))

    def get_token(self):
        try:
            token_data = f'client_id={self.dev_key}' \
                         f'&client_secret={self.dev_secret}' \
                         f'&scope=CompanyFile' \
                         f'&code={self.auth_code}' \
                         f'&redirect_uri={self.dev_redirect_uri}' \
                         f'&grant_type=authorization_code'
            headers = {'Content-Type': "application/x-www-form-urlencoded"}

            token = requests.request("POST",
                                     self.token_url,
                                     data=token_data,
                                     headers=headers)

            # Print access token Post response
            print(MSG_PREFIX + str(token.status_code))

            # Store returned JSON response.
            # Contains: access_token, token_type, scope, uid, username,
            # refresh_token, expires_in
            token_data = token.json()

            # Write new tokens to config and save config
            self.config.set('TOKENS', 'Access', token_data['access_token'])
            self.config.set('TOKENS', 'Refresh', token_data['refresh_token'])
            with open('settings.ini', 'w') as file:
                self.config.write(file)

            print(MSG_PREFIX + "Token Retrieved")

        except Exception as e:
            print(ERROR_PREFIX + str(e))

    def check_token(self):
        try:
            # Retrieve up to date access/refresh variables from config
            self.access_token = self.config.get('TOKENS', 'Access')
            self.refresh_token = self.config.get('TOKENS', 'Refresh')

            if self.access_token is None:
                print(MSG_PREFIX + 'No Access token')
            else:
                print(MSG_PREFIX + f'Access Token: {self.access_token}')
            if self.refresh_token is None:
                print(MSG_PREFIX + 'No Refresh Token')
            else:
                print(MSG_PREFIX + f'Refresh Token: {self.refresh_token}')
        except Exception as e:
            print(ERROR_PREFIX + str(e))

    def update_token(self):
        try:
            token_data = f'client_id={self.dev_key}' \
                         f'&client_secret={self.dev_secret}' \
                         f'&grant_type=refresh_token' \
                         f'&refresh_token={self.refresh_token}'
            headers = {'Content-Type': "application/x-www-form-urlencoded"}

            token = requests.request("POST",
                                     self.token_url,
                                     data=token_data,
                                     headers=headers)

            if DEBUG:
                # Prints access token Post response (200 = Success)
                print(token.status_code)

            # Store returned json response
            token_data = token.json()

            if DEBUG:
                print(token_data)

            # Write new tokens to config and save config
            self.config.set('TOKENS', 'Access', token_data['access_token'])
            self.config.set('TOKENS', 'Refresh', token_data['refresh_token'])
            with open('settings.ini', 'w') as file:
                self.config.write(file)

            # Updates access/refresh variables from config for requests
            self.access_token = self.config.get('TOKENS', 'Access')
            self.refresh_token = self.config.get('TOKENS', 'Refresh')

            print(MSG_PREFIX + "Tokens Successfully Updated")
        except Exception as e:
            print(ERROR_PREFIX + str(e))

    def myob_request(self, provided_request, pr_balance=0.0, pr_days_over=60):
        if provided_request == 'companyfile':
            request_choice = {'request_type': "GET",
                              'request_url':
                                  'https://api.myob.com/accountright',
                              'request_payload': {},
                              'request_parameters': {},
                              'request_headers':
                                  {'Authorization': 'Bearer '
                                                    + self.access_token,
                                   'x-myobapi-key': self.dev_key,
                                   'x-myobapi-version': 'v2',
                                   'Accept-Encoding': 'gzip,deflate'}}
        elif provided_request == 'overdue':

            def overdue_filter(balance=0.0, days_over=60):
                from datetime import datetime, timedelta, date
                """
                Retrieves and stores a MYOB JSON file with filters applied. 
                'Balance' is to be greater than the given value (default $0.00)
                and 'Date' must be prior to current date less days_over (default
                60 days).
                """
                d = datetime.today() - timedelta(days=days_over)

                uri = "https://ar2.api.myob.com/accountright/"
                file_id = ""
                invoices = "/Sale/Invoice"

                final_url = uri + file_id + invoices
                filters = f"?$filter=BalanceDueAmount gt {str(balance)}M and" \
                          f" Terms/DueDate le datetime'" \
                          f"{str(date(d.year, d.month, d.day).isoformat())}'"
                if DEBUG:
                    print(filters)
                    print(" Final URL: " + final_url + filters)
                return final_url + filters

            request_choice = {'request_type': "GET",
                              'request_url': overdue_filter(pr_balance,
                                                            pr_days_over),
                              'request_payload': {},
                              'request_parameters': {},
                              'request_headers':
                                  {'Authorization': 'Bearer '
                                                    + self.access_token,
                                   'x-myobapi-key': self.dev_key,
                                   'x-myobapi-version': 'v2',
                                   'Accept-Encoding': 'gzip,deflate'}}
        else:
            request_choice = {'request_type': "GET",
                              'request_url':
                                  'https://api.myob.com/accountright',
                              'request_payload': {},
                              'request_parameters': {},
                              'request_headers':
                                  {'Authorization': 'Bearer '
                                                    + self.access_token,
                                   'x-myobapi-key': self.dev_key,
                                   'x-myobapi-version': 'v2',
                                   'Accept-Encoding': 'gzip,deflate'}}

        try:
            if DEBUG:
                print(MSG_PREFIX + 'Retrieving Request...')
            response = \
                requests.request(request_choice['request_type'],
                                 request_choice['request_url'],
                                 headers=request_choice['request_headers'],
                                 params=request_choice['request_parameters'],
                                 data=request_choice['request_payload'])

            if DEBUG:
                print(MSG_PREFIX + 'Request data retrieved.')
            response = response.json()
            try:
                response_list = [response['Items']]
            except Exception as e:
                if DEBUG:
                    print(e)
                response_list = []
                pass

            # Checks if token is valid and current, or updates token.
            if 'Errors' in response:
                print(MSG_PREFIX + response['Errors'][0]['Name'])
                print(MSG_PREFIX + 'Attempting to refresh token and retry.')

                if self.retry_counter < 5:
                    try:
                        self.retry_counter += 1
                        self.update_token()
                        return self.myob_request(provided_request, pr_balance,
                                                 pr_days_over)
                    except Exception as e:
                        print(ERROR_PREFIX + str(e))
                        print(MSG_PREFIX + 'Re-authorisation may be required or '
                                           'there may be a network issue.')
                        self.get_user_authorisation()
                        self.get_token()
                        return self.myob_request(provided_request, pr_balance,
                                                 pr_days_over)
                elif self.retry_counter >= 5:
                    print(MSG_PREFIX + 'Re-authorisation may be required or '
                                       'there may be a network issue.')
                    self.get_user_authorisation()
                    self.check_token()
                    self.retry_counter = 0
                    return self.myob_request(provided_request, pr_balance,
                                             pr_days_over)
            elif 'NextPageLink' in response:
                try:
                    next_link = response['NextPageLink']
                    next_link_in_response = True

                    while next_link_in_response is True:
                        next_link_response = \
                            requests.request(request_choice['request_type'],
                                             url=next_link,
                                             headers=request_choice[
                                                 'request_headers'],
                                             params=request_choice[
                                                 'request_parameters'],
                                             data=request_choice[
                                                 'request_payload'])

                        next_link_response = next_link_response.json()
                        response_list.append(next_link_response['Items'])

                        if next_link_response['NextPageLink'] is not None:
                            next_link = next_link_response['NextPageLink']
                        else:
                            break

                except Exception as e:
                    print(ERROR_PREFIX + str(e))

            print('MYOB Authorisation request completed.')

            if DEBUG:
                print(len(response_list))

            return_list = []
            for r_list in response_list:
                for item in r_list:
                    return_list.append(item)

            if DEBUG:
                print(len(return_list))
                print(json.dumps(return_list, indent=4))

            return return_list

        except Exception as e:
            print(ERROR_PREFIX + str(e))


# Start Program:
if __name__ == "__main__":
    program_start(True)
