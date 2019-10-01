import { MsalAuthProvider, LoginType } from 'react-aad-msal';

const config = {
  auth: {
    authority: 'https://login.microsoftonline.com/' + process.env.REACT_APP_TENANT_ID,
    clientId: process.env.REACT_APP_CLIENT_ID || ''
  }
};

const authenticationParameters = {
    scopes: [
        'https://graph.microsoft.com/User.Read'
    ]
}

export const authProvider = new MsalAuthProvider(config, authenticationParameters, LoginType.Redirect)
