const { BetaAnalyticsDataClient } = require('@google-analytics/data');
const { GoogleAuth } = require('google-auth-library');



const serviceAccountKey = {
  type: process.env.GA_AUTHKEY_TYPE,
  project_id: process.env.GA_AUTHKEY_PROJECT_ID,
  private_key_id: process.env.GA_AUTHKEY_PRIVATE_KEY_ID,
  private_key: process.env.GA_AUTHKEY_PRIVATE_KEY.replace(/\\n/g, '\n'), 
  client_email: process.env.GA_AUTHKEY_CLIENT_EMAIL,
  client_id: process.env.GA_AUTHKEY_CLIENT_ID,
  auth_uri: process.env.GA_AUTHKEY_AUTH_URI,
  token_uri: process.env.GA_AUTHKEY_TOKEN_URI,
  auth_provider_x509_cert_url: process.env.GA_AUTHKEY_AUTH_PROVIDER_X509_CERT_URL,
  client_x509_cert_url: process.env.GA_AUTHKEY_CLIENT_X509_CERT_URL,
  universe_domain: process.env.GA_AUTHKEY_UNIVERSE_DOMAIN
};



    const auth = new GoogleAuth({
        credentials: serviceAccountKey,
        scopes: 'https://www.googleapis.com/auth/analytics.readonly',
    });

    const client = new BetaAnalyticsDataClient({
        auth
    });
             

const fetchGaData = async (dateRanges,dimensions,metrics,dimensionFilter=null,limit,proID) =>{

   if(typeof dimensionFilter =='number'){
     limit=dimensionFilter;
     dimensionFilter =null
   }
 
    try {
        
            const response = await client.runReport({
                property: `properties/${ proID || "353077807"}`,
                dateRanges: [dateRanges],
                dimensions: dimensions,     
                metrics: metrics,                  
                dimensionFilter: dimensionFilter,
                limit:limit,
            });

     return response;
    } catch (error) {
        console.error('Error fetching data:', error);
    }
}



module.exports = fetchGaData;