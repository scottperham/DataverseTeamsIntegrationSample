// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

$(async () => {

	const getTeamsToken = () => {
		return new Promise((resolve, reject) => {
			microsoftTeams.authentication.getAuthToken({
				successCallback: (token) => {
					resolve(token);
				},
				failureCallback: (reason) => {
					reject(reason);
				}
			})
		});
	}

	const initialiseTeams = () => {
		let rejectPromise = null;
		let timeout = null;

		const promise = new Promise((resolve, reject) => {
			rejectPromise = reject;
			microsoftTeams.initialize(() => {
				window.clearTimeout(timeout);
				resolve(true);
			}
			)
		});

		//At present this function will return that the page is not open in Teams
		//based on the MicrosoftTeams.initialize function timing out within 2 seconds
		timeout = window.setTimeout(() => {
			rejectPromise("Teams Initialise Timeout");
		}, 2000);

		return promise;
	}

	const getOnBehalfOfToken = async (teamsToken) => {
		const response = await fetch("/api/1.0/auth", {
			body: JSON.stringify({ Token: teamsToken }),
			method: "POST",
			headers: {
				'Content-Type': 'application/json'
			}
		});

		if (!response.ok) {
			return null;
		}

		const token = await response.text();
		return token;
	}

	const fetchAccountsFromDataverse = async (accessToken) => {
		const response = await fetch("https://itorgdev.crm11.dynamics.com/api/data/v9.1/accounts?$select=name&$expand=primarycontactid($select=contactid,fullname,emailaddress1)", {
			method: "GET",
			headers: {
				"Accept": "application/json",
				"Authorization": "bearer " + accessToken,
				"OData-MaxVersion": "4.0",
				"OData-Version": "4.0"
			}
		});

		if (!response.ok) {
			return null;
		}

		const result = await response.json();

		return result.value;
	}

	const createAccount = async (accountName) => {
		const response = await fetch("https://itorgdev.crm11.dynamics.com/api/data/v9.1/accounts", {
			method: "POST",
			body: JSON.stringify({name: accountName}),
			headers: {
				"Content-Type": "application/json",
				"Authorization": "bearer " + accessToken,
				"OData-MaxVersion": "4.0",
				"OData-Version": "4.0"
			}
		});

		if (!response.ok) {
			console.error(response.status);
			return false;
		}

		return true;
	}

	$("#newAccount").button().on("click", () => {
		dialog.dialog("open").position({
			my: "center",
			at: "center",
			of: window
		});
	});

	dialog = $("#dialog").dialog({
		autoOpen: false,
		height: 300,
		width: 350,
		modal: true,
		buttons: [
			{
				text: "Cancel",
				click: () => dialog.dialog("close")
			},
			{
				text: "Create account",
				click: async () => {
					if (await createAccount($("#accountName").val())) {
						dialog.dialog("close");
						await reloadData();
					}
				},
				class: "primary"
			}
		],
		dialogClass: "no-close"
	});

	const reloadData = async () => {

		const accounts = await fetchAccountsFromDataverse(accessToken);

		$("#loading").hide();

		$("#data tr.dynamic").remove();

		for (let i = 0; i < accounts.length; i++) {
			$("#data").html($("#data").html() + `<tr class='dynamic'><td>${accounts[i].name}</td><td>${accounts[i].primarycontactid?.fullname || "<em>None Set</em>"}</td><td>${accounts[i].primarycontactid?.emailaddress1 || "<em>None Set</em>"}</td></tr>`);
		}
	}

	try {
		await initialiseTeams();
	}
	catch {
		console.error("Not in Teams");
		return;
	}


	const teamsToken = await getTeamsToken();

	const accessToken = await getOnBehalfOfToken(teamsToken);

	await reloadData();
});