async function uploadFile(file) {
  try {
    ensureScope("files.readwrite");
    let options = {
      path: "/",
      fileName: file.name,
      rangeSize: 1024 * 1024, // must be a multiple of 320 KiB
    };
    const uploadTask = await MicrosoftGraph.OneDriveLargeFileUploadTask.create(
      graphClient,
      file,
      options
    );
    const response = await uploadTask.upload();
    console.log(`File ${response.name} of ${response.size} bytes uploaded`);
    return response;
  } catch (error) {
    console.error(error);
  }
}

function fileSelected(e) {
  displayUploadMessage(`Uploading ${e.files[0].name}...`);
  uploadFile(e.files[0]).then(({ responseBody: response }) => {
    displayUploadMessage(
      `File ${response.name} of ${response.size} bytes uploaded`
    );
    displayFiles();
  });
}

function displayUploadMessage(message) {
  const messageElement = document.getElementById("uploadMessage");
  messageElement.innerText = message;
}

async function downloadFile(file) {
  try {
    const response = await graphClient
      .api(`/me/drive/items/${file.id}`)
      .select("@microsoft.graph.downloadUrl")
      .get();
    const downloadUrl = response["@microsoft.graph.downloadUrl"];
    window.open(downloadUrl, "_self");
  } catch (error) {
    console.error(error);
  }
}

async function displayFiles() {
  const files = await getFiles();
  const ul = document.getElementById("downloadLinks");
  while (ul.firstChild) {
    ul.removeChild(ul.firstChild);
  }
  for (let file of files) {
    if (!file.folder && !file.package) {
      let a = document.createElement("a");
      a.href = "#";
      a.onclick = () => downloadFile(file);
      a.appendChild(document.createTextNode(file.name));
      let li = document.createElement("li");
      li.appendChild(a);
      ul.appendChild(li);
    }
  }
}

async function displayUI() {
  await signIn();

  // Display info from user profile
  const user = await getUser();
  var userName = document.getElementById("userName");
  userName.innerText = user.displayName;

  // Hide login button and initial UI
  var signInButton = document.getElementById("signin");
  signInButton.style = "display: none";
  var content = document.getElementById("content");
  content.style = "display: block";

  displayFiles();
}
