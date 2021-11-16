// Create an authentication provider
const authProvider = {
  getAccessToken: async () => {
    // Call getToken in auth.js
    return await getToken();
  },
};
// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });
//Get user info from Graph
async function getUser() {
  ensureScope("user.read");
  return await graphClient.api("/me").select("id,displayName").get();
}

async function getFiles() {
  ensureScope("files.read");
  try {
    const response = await graphClient
      .api("/me/drive/root/children")
      .select("id,name,folder,package")
      .get();
    return response.value;
  } catch (error) {
    console.error(error);
  }
}
