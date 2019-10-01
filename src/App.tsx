import React from 'react';
import logo from './logo.svg';
import './App.css';

import { MsalAuthProvider } from 'react-aad-msal';
import { AzureAD, AuthenticationState, IAzureADFunctionProps } from 'react-aad-msal';
import { authProvider } from './authProvider';

import { ServiceBus } from './serviceBus';

interface IAppProps {
}

interface IAppState {
  result: string
}

class App extends React.Component<IAppProps, IAppState> {
  serviceBus: ServiceBus;
  provider: MsalAuthProvider;

  constructor(props: IAppProps) {
    super(props);
    this.state = {
      result: ''
    };
    this.serviceBus = new ServiceBus();
    this.provider = authProvider;
  }

  async componentDidMount() {
    if (authProvider.authenticationState === AuthenticationState.Unauthenticated) {
      return;
    }
    this.serviceBus.on('result', (result) => {
      this.setState({
        result: result
      })
    });
    await this.serviceBus.initialize(this.provider);
  }

  async componentWillUnmount() {
    await this.serviceBus.uninitialize();
  }

  renderButtons() {
    if (authProvider.authenticationState === AuthenticationState.Unauthenticated) {
      return;
    }
    return (
      <div>
        <p>
          <button onClick={() => { this.serviceBus.getSubscriptions() }}>Get Subscriptions</button>
        </p>
        <p>
          <button onClick={() => { this.serviceBus.createSubscription() }}>Create Subscription</button>
        </p>
        <p>
          <button onClick={() => { this.serviceBus.send() }}>Send</button>
        </p>
        <p>
          {this.state.result}
        </p>
      </div>
    )
  }

  render() {
    return (
      <div className="App">
        <header className="App-header">
          <img src={logo} className="App-logo" alt="logo" />
          <AzureAD provider={authProvider}>
            {
              ({ login, logout, authenticationState, accountInfo }: IAzureADFunctionProps) => {
                if (authenticationState === AuthenticationState.Authenticated) {
                  return (
                    <>
                      <span>Welcome, {accountInfo && accountInfo.account.name}!</span>
                      <button className="App-button" onClick={logout}>Logout</button>
                    </>
                  );
                } else if (authenticationState === AuthenticationState.Unauthenticated) {
                  return (
                    <>
                      <span>Hey stranger, you look new!</span>
                      <button className="App-button" onClick={login}>Login</button>
                    </>
                  );
                }
              }
            }
          </AzureAD>
        </header>
        {this.renderButtons()}
      </div>
    );
  }
}

export default App;
