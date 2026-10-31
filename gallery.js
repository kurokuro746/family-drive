console.log("script start5");

const msalConfig = {
    auth: {
        clientId: "815ff81e-8462-4eb0-a444-5a63070a1425",
        authority: "https://login.microsoftonline.com/consumers",
        redirectUri: "https://kurokuro746.github.io/family-drive/"
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
let folderPath = [];

        (async () => {

            await msalInstance.initialize();

            const response =
                await msalInstance.handleRedirectPromise();

            if (response) {
                showUser(response.account);
                await listFiles();
            } else {
                const accounts =
                    msalInstance.getAllAccounts();
                if (accounts.length > 0) {
                    showUser(accounts[0]);
                    await listFiles();
                }
        }
        })();
       
        async function signIn() {

            try {
                await msalInstance.loginRedirect({
                    scopes: [
                        "User.Read",
                        "Files.Read",
                        "Files.ReadWrite"
                    ]
                });
            } catch (error) {
                console.error(error);
            }
        }

        async function signOut() {
            const account = msalInstance.getAllAccounts()[0];
            if (account) {
                await msalInstance.logoutPopup({
                    account: account
                });
            }
            location.reload();

        }
        function displayFiles(data, currentDriveId = null) {
            let html = "<h2>OneDriveファイル一覧</h2><ul>";
            for (const item of data.value) {

                if (item.folder || item.remoteItem?.folder) {
                    const folderId = item.remoteItem?.id || item.id;
                    const driveId = item.remoteItem?.parentReference?.driveId || currentDriveId || null;
                    html += `
                        <li>
                            📁
                            <a href="#"
                               onclick="openFolder('${folderId}','${item.name}','${driveId}')">
                                ${item.name}
                            </a>
                        </li>
                    `;
                } else if (item.image || item.video) {
                    const thumbnailUrl = item.thumbnails?.[0]?.medium?.url;
                    const isVideo = !!item.video;
  
                    html += `
                        <div style="display:inline-block;margin:10px;position:relative;">
                            <img
                                src="${thumbnailUrl}"
                                onclick="viewFile(
                                    '${item["@microsoft.graph.downloadUrl"]}',
                                    '${item.image ? "image" : "video"}'
                                )"
                                style="
                                    width:150px;
                                    height:150px;
                                    object-fit:cover;
                                    cursor:pointer;
                                "
                            >
                            ${isVideo ? `
                                <div style="
                                    position:absolute;
                                    right:5px;
                                    bottom:5px;
                                    background:rgba(0,0,0,0.7);
                                    color:white;
                                    padding:2px 6px;
                                    border-radius:4px;
                                    font-size:20px;
                                ">
                                    ▶
                                </div>
                            ` : ""}
                        </div>
                    `;
                }
            }
            html += "</ul>";
            document.getElementById("fileList").innerHTML =
                html;
        }

        async function listFiles() {
            folderPath = [
                {
                    id: null,
                    name: "ルート",
                    driveId: null
                }
            ];
            renderBreadcrumb();
            try {
                const account =
                    msalInstance.getAllAccounts()[0];
                const tokenResponse =
                    await msalInstance.acquireTokenSilent({
                        scopes: [
                            "Files.Read"
                        ],
                        account: account
                    });
                const response =
                    await fetch(
                        "https://graph.microsoft.com/v1.0/me/drive/root/children",
                        {
                            headers: {
                                Authorization:
                                    "Bearer " +
                                    tokenResponse.accessToken
                            }
                        }
                    );
                const data = await response.json();
                const shortcutFolder = data.value.find(item => item.name === "ショートカット");
                if (shortcutFolder) {
                    folderPath = [];
                    openFolder(shortcutFolder.id, shortcutFolder.name, null);
                    return;
                }
                displayFiles(data);
            } catch(error) {
                console.error(error);
            }
        }
        async function openFolder(folderId, folderName, driveId) {
            const currentDriveId = driveId;
            folderPath.push({
                id: folderId,
                name: folderName,
                driveId: driveId
            });
            renderBreadcrumb();
            try {
                const account =
                    msalInstance.getAllAccounts()[0];
                const tokenResponse =
                    await msalInstance.acquireTokenSilent({
                        scopes: ["Files.Read"],
                        account: account
                    });
                let url;
                if (driveId && driveId !== "null") {
                    url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}/children?$expand=thumbnails`;
                } else {
                    url = `https://graph.microsoft.com/v1.0/me/drive/items/${folderId}/children?$expand=thumbnails`;
                }
                const response =
                    await fetch(url,
                        {
                            headers: {
                                Authorization:
                                    "Bearer " +
                                    tokenResponse.accessToken
                            }
                        }
                    );
                const data = await response.json();
                displayFiles(data, currentDriveId);
            } catch(error) {
                console.error(error);
            }
        }
        async function openFolderById(folderId, driveId) {
            try {
                const account =
                    msalInstance.getAllAccounts()[0];
                const tokenResponse =
                    await msalInstance.acquireTokenSilent({
                        scopes: ["Files.Read"],
                        account: account
                    });
                let url;
                if (driveId && driveId !== "null") {
                    url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}/children?$expand=thumbnails`;
                } else {
                    url = `https://graph.microsoft.com/v1.0/me/drive/items/${folderId}/children?$expand=thumbnails`;
                }
                const response =
                    await fetch(url,
                        {
                            headers: {
                                Authorization:
                                    "Bearer " +
                                    tokenResponse.accessToken
                            }
                        }
                    );
                const data = await response.json();
                displayFiles(data, driveId);
            } catch(error) {
                console.error(error);
            }
        }
        function renderBreadcrumb() {
            let html = "";
            for (let i = 0; i < folderPath.length; i++) {
                const displayName = i === 0? "ルート": folderPath[i].name;
                if (i > 0) {
                    html += " &gt; ";
                }
                html += `
                    <a href="#"
                    onclick="navigateToBreadcrumb(${i})">
                        ${displayName}
                    </a>
                `;
            }
            document.getElementById("breadcrumb").innerHTML =
                html;
        }

        async function navigateToBreadcrumb(index) {
            folderPath = folderPath.slice(0, index + 1);
            const targetFolder =
                folderPath[folderPath.length - 1];

            if (targetFolder.id === null) {
                await listFiles();
                return;
            }
            await openFolderById(targetFolder.id, targetFolder.driveId);
            renderBreadcrumb();
        }
        async function goBack() {
            if (folderPath.length <= 1) {
                return;
            }
            folderPath.pop();
            renderBreadcrumb();
            const parentFolder =
                folderPath[folderPath.length - 1];
            if (parentFolder.id === null) {
                await listFiles();
                return;
            }
            await openFolderById(parentFolder.id, parentFolder.driveId);
        }
        function openFile(url) {
            window.open(url, "_blank");
        }

        function viewFile(url, type) {
            const modal = document.getElementById("modal");
            if (type === "image") {
                modal.innerHTML = `
                    <div
                        onclick="closeModal()"
                        style="
                            position:absolute;
                            top:20px;
                            right:30px;
                            color:white;
                            font-size:40px;
                            cursor:pointer;
                        "
                    >
                        ×
                    </div>

                    <img
                        src="${url}"
                        style="
                            max-width:95%;
                            max-height:95%;
                            margin-top:20px;
                        "
                    >
                `;
            } else {
                modal.innerHTML = `
                    <div
                        onclick="closeModal()"
                        style="
                            position:absolute;
                            top:20px;
                            right:30px;
                            color:white;
                            font-size:40px;
                            cursor:pointer;
                        "
                    >
                        ×
                    </div>
                    <video
                        controls
                        autoplay
                        style="
                            max-width:95%;
                            max-height:95%;
                            margin-top:20px;
                        "
                    >
                        <source src="${url}">
                    </video>
                `;
            }
            modal.style.display = "block";
        }
        function closeModal() {
            const modal = document.getElementById("modal");

            modal.style.display = "none";
            modal.innerHTML = "";
        }

        function showUser(account) {

            document.getElementById("userInfo").innerHTML =
                `
                <p>ログイン中</p>
                <p>${account.username}</p>
                `;

            document.getElementById("loginButton").style.display =
                "none";

            document.getElementById("logoutButton").style.display =
                "inline";

        }

        const accounts = msalInstance.getAllAccounts();

        if (accounts.length > 0) {

            showUser(accounts[0]);

        }

        document
            .getElementById("loginButton")
            .addEventListener("click", signIn);
        document
            .getElementById("logoutButton")
            .addEventListener("click", signOut);
        document
            .getElementById("backButton")
            .addEventListener("click", goBack);
