
require('isomorphic-fetch');
const azure=require('@azure/identity');
const graph=require('@microsoft/microsoft-graph-client');
const authProviders=require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

let _settings=undefined;
let _deviceCodeCredential=undefined;
let _userClient=undefined;

function initializeGraphForUserAuth(settings,deviceCodePrompt){
    if(!settings){
        throw new Error('Settings cannot be undefined');
    }

    _settings=settings;
    _deviceCodeCredential=new azure.DeviceCodeCredential({
        clientId:settings.clientId,
        tenantId:settings.authTenant,
        userPromptCallback:deviceCodePrompt
    });

    const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
        _deviceCodeCredential, {
          scopes: settings.graphUserScopes
        });

    _userClient=graph.Client.initWithMiddleware({
        authProvider:authProvider
    });
    
}
module.exports.initializeGraphForUserAuth=initializeGraphForUserAuth;


async function getUserTokenAsync(){
    if(!_deviceCodeCredential){
        throw new Error('Graph no ha sido inicializado por el user auth');
    }

    if(!_settings?.graphUserScopes){
        throw new Error('El alcance de la configuracion no esta definido');
    }

    const response = await _deviceCodeCredential.getToken(_settings?.graphUserScopes);
    return response.token;
}
module.exports.getUserTokenAsync=getUserTokenAsync;

async function getUserAsync(){
    if(!_userClient){
        throw new Error('Graph no fue inicializado por el user auth');
    }
    return _userClient.api('/me')
    .select(['displayName','mail','userPrincipalName'])
    .get();
}
module.exports.getUserAsync=getUserAsync;




