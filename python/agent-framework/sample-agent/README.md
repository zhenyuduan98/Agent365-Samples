# Agent Framework Sample Agent - Python

This sample demonstrates how to build an agent using Agent Framework in Python with the Microsoft Agent 365 SDK. It covers:

- **Observability**: End-to-end tracing, caching, and monitoring for agent applications
- **Notifications**: Services and models for managing user notifications
- **Tools**: Model Context Protocol tools for building advanced agent solutions
- **Hosting Patterns**: Hosting with Microsoft 365 Agents SDK

This sample uses the [Microsoft Agent 365 SDK for Python](https://github.com/microsoft/Agent365-python).

For comprehensive documentation and guidance on building agents with the Microsoft Agent 365 SDK, including how to add tooling, observability, and notifications, visit the [Microsoft Agent 365 Developer Documentation](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/).

## Prerequisites

- Python 3.x
- Microsoft Agent 365 SDK
- Agent Framework (agent-framework-azure-ai)
- Azure/OpenAI API credentials

## Working with User Identity

On every incoming message, the A365 platform populates `activity.from_property` with basic user
information — always available with no API calls or token acquisition:

| Field | Description |
|---|---|
| `activity.from_property.id` | Channel-specific user ID (e.g., `29:1AbcXyz...` in Teams) |
| `activity.from_property.name` | Display name as known to the channel |
| `activity.from_property.aad_object_id` | Azure AD Object ID — use this to call Microsoft Graph |

The sample logs these fields at the start of every message turn and injects the display name
into the LLM system instructions for personalized responses.

## Handling Agent Install and Uninstall

When a user installs (hires) or uninstalls (removes) the agent, the A365 platform sends an `InstallationUpdate` activity — also referred to as the `agentInstanceCreated` event. The sample handles this in `on_installation_update` in `host_agent_server.py`:

| Action | Description |
|---|---|
| `add` | Agent was installed — send a welcome message |
| `remove` | Agent was uninstalled — send a farewell message |

```python
if action == "add":
    await context.send_activity("Thank you for hiring me! Looking forward to assisting you in your professional journey!")
elif action == "remove":
    await context.send_activity("Thank you for your time, I enjoyed working with you.")
```

To test with Agents Playground, use **Mock an Activity → Install application** to send a simulated `installationUpdate` activity.

## Entra ID App Registration (Step-by-Step)

Before running this agent, you need to register an application in Microsoft Entra ID (Azure AD). This app provides the `CLIENT_ID` and `CLIENT_SECRET` used for authentication against the Microsoft 365 Agents platform and Microsoft Graph.

### Step 1: Create the App Registration

1. Go to [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **+ New registration**
2. Fill in the details:
   - **Name**: e.g., `Agent365-CLI-Client` (or any descriptive name)
   - **Supported account types**: Select **Accounts in this organizational directory only** (Single tenant)
   - **Redirect URI**: Select **Public client/native (mobile & desktop)** and enter `http://localhost`
3. Click **Register**
4. On the **Overview** page, copy the following values — you will need them for `.env`:
   - **Application (client) ID** → `CLIENT_ID`
   - **Directory (tenant) ID** → `TENANT_ID`

### Step 2: Create a Client Secret

1. In your app registration, go to **Certificates & secrets** → **Client secrets** → **+ New client secret**
2. Enter a description (e.g., `agent-dev-secret`) and choose an expiration period
3. Click **Add**
4. **Immediately copy the secret Value** (not the Secret ID) → `CLIENT_SECRET`
   > ⚠️ The secret value is only shown once. If you lose it, you must create a new one.

### Step 3: Configure API Permissions (Delegated)

This agent needs delegated permissions to access Microsoft Graph and MCP tool servers on behalf of the signed-in user.

1. Go to **API permissions** → **+ Add a permission**

#### 3a. Microsoft Graph Permissions

Select **Microsoft Graph** → **Delegated permissions**, then add:

| Permission | Purpose |
|---|---|
| `Mail.Read` | Read user's mailbox (for MCP MailTools) |
| `Mail.Send` | Send mail on behalf of the user |
| `User.Read` | Sign in and read user profile |

#### 3b. MCP Server Permissions (Agent 365 Platform)

If your agent uses MCP tool servers (e.g., `mcp_MailTools`), you also need to add permissions for the MCP server API:

1. Click **+ Add a permission** → **APIs my organization uses**
2. Search for the MCP server application by its audience (e.g., `05879165-0320-489e-b644-f72b33f3edf0` — this is the MCP MailTools resource app)
3. Select **Delegated permissions** → add:

| Permission | Purpose |
|---|---|
| `McpServers.Mail.All` | Access MCP Mail tool server |
| `McpServersMetadata.Read.All` | Read MCP server metadata |

> 💡 These scopes correspond to the `scope` and `audience` fields in `ToolingManifest.json`.

#### 3c. Grant Admin Consent

After adding all permissions, click **Grant admin consent for [your tenant]** (requires Global Admin or Privileged Role Admin).

Verify that all permissions show a green checkmark ✅ under the **Status** column.

### Step 4: Configure the Service Connection

For the M365 Agents SDK to authenticate outgoing API calls, this sample uses a **Service Connection** pattern via environment variables. Add the following to your `.env` file:

```env
# App Registration
CLIENT_ID=<your-application-client-id>
CLIENT_SECRET=<your-client-secret-value>
TENANT_ID=<your-directory-tenant-id>

# Service Connection (used by M365 Agents SDK)
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID=<same-client-id>
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET=<same-client-secret>
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID=<same-tenant-id>
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__SCOPES=https://graph.microsoft.com/.default
CONNECTIONSMAP_0_SERVICEURL=*
CONNECTIONSMAP_0_CONNECTION=SERVICE_CONNECTION
```

> The `CONNECTIONS__*` variables map to the M365 Agents SDK connection configuration. Setting `CONNECTIONSMAP_0_SERVICEURL=*` means this connection is used for all outbound service URLs.

### Step 5: (Optional) MCP Bearer Token

If using MCP tool servers, the agent needs a user-delegated token with MCP scopes. This token is passed to the MCP server for on-behalf-of access:

```env
MCP_BEARER_TOKEN=<user-delegated-jwt-with-mcp-scopes>
```

To obtain this token:
1. Use the [MSAL](https://learn.microsoft.com/en-us/entra/identity-platform/msal-overview) library or [Microsoft Authentication Manager (WAM)](https://learn.microsoft.com/en-us/entra/identity-platform/scenario-desktop-acquire-token-wam) on a Windows machine to sign in interactively
2. Request scopes: `api://05879165-0320-489e-b644-f72b33f3edf0/McpServers.Mail.All`
3. Copy the resulting access token into `MCP_BEARER_TOKEN`

> ⚠️ This token expires after ~1 hour. For development/demo purposes, you need to refresh it manually.

### Summary of `.env` Variables

| Variable | Description |
|---|---|
| `CLIENT_ID` | Entra App Registration client ID |
| `CLIENT_SECRET` | Entra App Registration client secret |
| `TENANT_ID` | Entra Directory (tenant) ID |
| `AZURE_OPENAI_API_KEY` | Azure OpenAI API key |
| `AZURE_OPENAI_ENDPOINT` | Azure OpenAI endpoint URL |
| `AZURE_OPENAI_DEPLOYMENT` | Azure OpenAI model deployment name |
| `AZURE_OPENAI_API_VERSION` | Azure OpenAI API version |
| `PORT` | Agent server listening port (default: `3978`) |
| `MCP_BEARER_TOKEN` | User-delegated JWT for MCP server access |
| `USE_AGENTIC_AUTH` | Enable agentic (Agent Instance) auth (`true`/`false`) |
| `LOG_LEVEL` | Logging level (`INFO`, `DEBUG`, etc.) |

### Troubleshooting

| Issue | Solution |
|---|---|
| `AADSTS65001: The user or administrator has not consented` | Go to API permissions and click **Grant admin consent** |
| `AADSTS7000215: Invalid client secret` | Regenerate the secret; make sure you copied the **Value**, not the **Secret ID** |
| `Directory.AccessAsUser.All` causing blueprint creation failure | This permission may be injected by WAM; clear token cache and re-authenticate |
| MCP token expired | Refresh the `MCP_BEARER_TOKEN` by re-authenticating with MSAL/WAM |
| `401 Unauthorized` in Teams | Agent needs an Agent Instance with proper credentials (anonymous mode won't work) |

---

## Running the Agent

To set up and test this agent, refer to the [Configure Agent Testing](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/testing?tabs=python) guide for complete instructions.

For a detailed explanation of the agent code and implementation, see the [Agent Code Walkthrough](AGENT-CODE-WALKTHROUGH.md).

## Support

For issues, questions, or feedback:

- **Issues**: Please file issues in the [GitHub Issues](https://github.com/microsoft/Agent365-python/issues) section
- **Documentation**: See the [Microsoft Agents 365 Developer documentation](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/)
- **Security**: For security issues, please see [SECURITY.md](SECURITY.md)

## Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit <https://cla.opensource.microsoft.com>.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Additional Resources

- [Microsoft Agent 365 SDK - Python repository](https://github.com/microsoft/Agent365-python)
- [Microsoft 365 Agents SDK - Python repository](https://github.com/Microsoft/Agents-for-python)
- [Agent Framework documentation](https://github.com/microsoft/Agent365-python/tree/main/packages/agent-framework)
- [Python API documentation](https://learn.microsoft.com/python/api/?view=m365-agents-sdk&preserve-view=true)

## Trademarks

*Microsoft, Windows, Microsoft Azure and/or other Microsoft products and services referenced in the documentation may be either trademarks or registered trademarks of Microsoft in the United States and/or other countries. The licenses for this project do not grant you rights to use any Microsoft names, logos, or trademarks. Microsoft's general trademark guidelines can be found at http://go.microsoft.com/fwlink/?LinkID=254653.*

## License

Copyright (c) Microsoft Corporation. All rights reserved.

Licensed under the MIT License - see the [LICENSE](LICENSE.md) file for details.