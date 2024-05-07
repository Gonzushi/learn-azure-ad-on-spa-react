import './App.css';
import { PageLayout } from './components/PageLayout';
import { ProfileData } from './components/ProfileData';

import React, { useState } from 'react';

import { loginRequest } from './authConfig';
import { callMsGraph } from './graph';

import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';

import Button from 'react-bootstrap/Button';


const ProfileContent = () => {
  const { instance, accounts } = useMsal();
  const [graphData, setGraphData] = useState(null);

  function RequestProfileData() {
      // Silently acquires an access token which is then attached to a request for MS Graph data
      instance
          .acquireTokenSilent({
              ...loginRequest,
              account: accounts[0],
          })
          .then((response) => {
            console.log(response.accessToken)
              callMsGraph(response.accessToken).then((response) => setGraphData(response));
          });
  }

  return (
      <>
          <h5 className="card-title">Welcome {accounts[0].name}</h5>
          <br/>
          {graphData ? (
              <ProfileData graphData={graphData} />
          ) : (
              <Button variant="secondary" onClick={RequestProfileData}>
                  Request Profile Information
              </Button>
          )}
      </>
  );
};

const MainContent = () => {
  return (
      <div className="App">
          <AuthenticatedTemplate>
              <ProfileContent />
          </AuthenticatedTemplate>

          <UnauthenticatedTemplate>
              <h5>
                  <center>
                      Please sign-in to see your profile information.
                  </center>
              </h5>
          </UnauthenticatedTemplate>
      </div>
  );
};

function App() {
  return (
    <PageLayout>
      <MainContent/>
    </PageLayout>
  );
}

export default App;
