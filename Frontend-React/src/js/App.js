import React, { useState, useEffect } from 'react';

//https://wangchujiang.com/react-monacoeditor/
import MonacoEditor from '@uiw/react-monacoeditor';

import { AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
//https://www.npmjs.com/package/@azure/msal-browser
import { InteractionType, PublicClientApplication, EventType, LogLevel, EventMessage} from '@azure/msal-browser';
import { useMsal,MsalProvider  } from '@azure/msal-react';

//https://material-ui.com/components/autocomplete/
import {TextField, Button} from '@material-ui/core';
import {Autocomplete} from '@material-ui/lab';

import Config from '../Config';

import {callMSGraph, parseJwt} from './main'
import '../css/main.css';

const msalConfig = {
  auth: {
      clientId: "8a792f49-ae0d-4b9b-92d2-614fcba43bea",
      authority: "https://login.microsoftonline.com/72f988bf-86f1-41af-91ab-2d7cd011db47",
      redirectUri: "http://localhost:3000"
  },
  cache: {
      cacheLocation: "sessionStorage", // This configures where your cache will be stored
      storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
  },
  system: {
      loggerOptions: {
          loggerCallback: (level, message, containsPii) => {
              if (containsPii) {	
                  return;	
              }	
              switch (level) {	
                  case LogLevel.Error:	
                      console.error(message);	
                      return;	
                  case LogLevel.Info:	
                      console.info(message);	
                      return;	
                  case LogLevel.Verbose:	
                      console.debug(message);	
                      return;	
                  case LogLevel.Warning:	
                      console.warn(message);	
                      return;	
              }
          }
      }
  }
};

function App() {

  const loginRequest = {
    scopes: ["User.Read"]
  };

  let [msalInstance, setMsalInstance] = useState(null);
  let [lastRequest, setLastRequest] = useState({});

  ///////////////////////////////////////////////////////////////////
  // Set Graph Call Type
  let [graph, setGraph] = useState("Microsoft Graph");
  let changeGraph = () => {
    if(graph ==="Microsoft Graph")
      setGraph("AAD Graph")
    else
      setGraph("Microsoft Graph")
  }
  
  ///////////////////////////////////////////////////////////////////
  // Set Account Information
  let [account, setAccount] = useState("Sign In");
  let [accountInfo, setAccountInfo] = useState(null);
  let [accountId, setAccountId] = useState("");
  let changeSignIn = () => {
    if(account ==="Sign In"){
      msalInstance.handleRedirectPromise().then(handleResponse).catch((error) => {
        console.log(error);
      });
    }else
      signOut()
  }
  let handleResponse = async (resp) => {
    console.log(resp)
    if (resp !== null) {
      setAccountId(resp.account.homeAccountId);
      setAccessToken(resp.accessToken)
      let decoded = await parseJwt(resp.accessToken);
      let stringDecoded = await JSON.stringify(decoded)
      stringDecoded = await stringDecoded.replaceAll(",", ",\n");
      setDecodedToken(stringDecoded)
      setTokenAquired(true)
      setAccountInfo(resp.account)
      setLastRequest(resp)
    }else{
      const currentAccounts = msalInstance.getAllAccounts();
      console.log(currentAccounts)
      if (!currentAccounts || currentAccounts.length < 1) {
        signIn("loginRedirect");
        setAccount("Sign Out")
      } else if (currentAccounts.length > 1) {
          // Add choose account code here
      } else if (currentAccounts.length === 1) {
          setAccountId(currentAccounts[0].homeAccountId);
          setAccountInfo(currentAccounts[0])
          setAccount("Sign Out")
      }
    }
  }
  async function signIn(method) {
    //signInType = isIE ? "loginRedirect" : method;
    let signInType = method
    // if (signInType === "loginPopup") {
      return msalInstance.loginPopup(loginRequest).then(handleResponse).catch(function (error) {
          console.log(error);
      });
    // } else if (signInType === "loginRedirect") {
    //     return msalInstance.loginRedirect(loginRequest);
    // }
  }
  function signOut() {
    const logoutRequest = {
        account: msalInstance.getAccountByHomeId(accountId)
    };
    msalInstance.logoutRedirect(logoutRequest);
    setAccount("Sign In")
    setAccessToken("")
  }

  ///////////////////////////////////////////////////////////////////
  // Set AccessToken
  let [tokenAquired, setTokenAquired] = useState(false);
  let [accessToken, setAccessToken] = useState("");
  let [decodedToken, setDecodedToken] = useState("");
  async function getTokenPopup(request, account) {
    if(!accountInfo){
      alert("Login First")
      return null;
    }
    request.account = account;
    let results = await msalInstance.acquireTokenSilent(request).catch(async (error) => {
        console.log("silent token acquisition fails.");
        if (error instanceof msalInstance.InteractionRequiredAuthError) {
            console.log("acquiring token using popup");
            return msalInstance.acquireTokenPopup(request).catch(error => {
                console.error(error);
            });
        } else {
            console.error(error);
        }
    });
    console.log("getTokenPopup :", results)
    let decoded = await parseJwt(results.accessToken);
    let stringDecoded = await JSON.stringify(decoded)
    stringDecoded = await stringDecoded.replaceAll(",", ",\n");

    setDecodedToken(stringDecoded)
    setAccessToken(results.accessToken)
    setLastRequest(results)
    return
  }

  ///////////////////////////////////////////////////////////////////
  // Graph Calls
  let [selectedUri, setSelectedUri] = useState("");
  let [selectedVersion, setSelectedVersion] = useState("");
  let [selectedEndpoint, setSelectedEndpoint] = useState("");
  let [graphResults, setGraphResults] = useState("");
  function MakeCall()
  {
    // if(accessToken)
    //   callMSGraph(selectedUri, selectedVersion, selectedEndpoint, accessToken, SetResults);
    // Simple POST request with a JSON body using fetch
    const requestOptions = {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ 
        uri : "https://graph.microsoft.com/",
        version : "v1.0",
        endpoint : "me"
       })
    };
    fetch('http://localhost:3000/GraphCall', requestOptions)
      .then(response => console.log(response.json()))
      //.then(data => this.setState({ postId: data.id }));

  }

  async function SetResults (response, endpoint) {
    console.log(response,endpoint)
    let stringResponse = await JSON.stringify(response)
    stringResponse = await stringResponse.replaceAll(",", ",\n");
    stringResponse = await stringResponse.replaceAll("[", "\n[\n");
    stringResponse = await stringResponse.replaceAll("]", "\n]\n");
    stringResponse = await stringResponse.replaceAll("{", "\n{\n");
    stringResponse = await stringResponse.replaceAll("}", "\n}\n");
    setGraphResults(stringResponse)
  }

  useEffect(() => {
    setMsalInstance(new PublicClientApplication(msalConfig)); 
  },[]);

  return (
    <div className="">

      {/* Navbar */}
      <div className="nav row" >
        <div className="col-8"> Graph Helper </div>
        <div className="col-4 btn-group" role="group" aria-label="...">
          <button type="button" className="btn" onClick={changeGraph}>{graph}</button>
          <button type="button" className="btn" onClick={changeSignIn}>{account}</button>
          <button type="button" className="btn" onClick={()=>console.log(getTokenPopup(loginRequest, accountInfo))} >Get {tokenAquired? "New": ""} Token</button>
        </div>
      </div>

      <div className="row">
        {/* Graph Section */}
        <div className="col-8 row" style={{background: "#f4f4fa"}}>

          <div className="col-4">
            <Autocomplete 
              style={{width:"100%"}}
              value={selectedUri}
              onChange={(event, newValue) => { setSelectedUri(newValue);}}
              onInputChange={(event, newInputValue) => {setSelectedUri(newInputValue);}}
              options={Config.uri.map((option) => option)}
              renderInput={(params) => ( <TextField {...params} label="Graph Uri"/>)}
            />
          </div>
          <div className="col-2">
            <Autocomplete
              style={{width:"100%"}}
              value={selectedVersion}
              onChange={(event, newValue) => { setSelectedVersion(newValue);}}
              onInputChange={(event, newInputValue) => {setSelectedVersion(newInputValue);}}
              options={Config.versions.map((option) => option)}
              renderInput={(params) => ( <TextField {...params} label="Version"/>)}
            />
          </div>
          <div className="col-4">
            <Autocomplete
                style={{width:"100%"}}
                value={selectedEndpoint}
                onChange={(event, newValue) => { setSelectedEndpoint(newValue);}}
                onInputChange={(event, newInputValue) => {setSelectedEndpoint(newInputValue);}}
                options={Config.endpoints.map((option) => option)}
                renderInput={(params) => ( <TextField {...params} label="Endpoints"/>)}
            />
          </div>
          <div className="col-2">
            <Button 
              onClick={MakeCall}
              variant="contained" 
              style={{marginTop: 10, width: "100%"}}>
                Query
            </Button>
          </div>
        
          <div className="col-12">
            <MonacoEditor
                height="325px"
                language="json"
                value={graphResults}
                options={{
                  selectOnLineNumbers: true,
                  roundedSelection: false,
                  cursorStyle: 'line',
                  automaticLayout: true,
                  theme: 'vs-dark',
                }}
              />
          </div>
        </div>

        {/* Token Info */}
        <div className="col-4" style={{background: "#eee"}}>

          <MonacoEditor
            height="400px"
            language="json"
            formatOnPaste="true"
            value={String(decodedToken)}
            options={{
              selectOnLineNumbers: true,
              roundedSelection: false,
              wordWrap: "on",
              autoIndent: true,
              cursorStyle: 'line',
              readOnly: true,
              automaticLayout: true,
              theme: 'vs-dark',
              tabSize: 2,
              autoIndent: true
            }}
          />
          
        </div>
      </div>

      <p>msalInstance : {msalInstance? msalInstance.length : "0"}</p>

      <p>accessToken : {accessToken}</p>

      <p>accountId : {accountId}</p>

    </div>
  );
}

export default App;

