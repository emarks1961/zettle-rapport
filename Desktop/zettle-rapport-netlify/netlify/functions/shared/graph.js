'use strict';
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.GRAPH_TENANT_ID,
    process.env.GRAPH_CLIENT_ID,
    process.env.GRAPH_CLIENT_SECRET
  );
  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const token = await credential.getToken('https://graph.microsoft.com/.default');
        return token.token;
      }
    }
  });
}

module.exports = { getGraphClient };
