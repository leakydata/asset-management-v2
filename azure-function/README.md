# Cat Asset Management — read-only broker (Azure Function)

A small HTTP Function that holds the Cat API credentials server-side and exposes
a single read-only `/api/search` endpoint. Excel (and anything else) calls this
instead of Cat, so the client secret is never distributed.

`function_app.py` reuses the same OAuth client-credentials + search logic as the
MCP server. Only search is exposed — there are no write routes.

---

## Phase 0 — Install the tooling (one time)

```powershell
winget install Microsoft.AzureCLI
winget install Microsoft.Azure.FunctionsCoreTools
npm install -g azurite          # local storage emulator (or use the VS Code Azurite ext)
```

Azure Functions does **not** support Python 3.14. Use 3.11 for this project:

```powershell
uv python install 3.11
```

Then sign in:

```powershell
az login
```

---

## Phase 1 — Run it locally

```powershell
# from azure-function\
uv venv --python 3.11
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

copy local.settings.json.example local.settings.json
# edit local.settings.json: fill CAT_CLIENT_ID, CAT_CLIENT_SECRET, CAT_SCOPE
```

Start the storage emulator (separate terminal) and the function:

```powershell
azurite            # terminal 1
func start         # terminal 2  (from azure-function\)
```

Test it (no key needed locally):

```powershell
curl "http://localhost:7071/api/search?serial=9303"
```

You should get the same JSON the Cat API returns.

---

## Phase 2 — Create the Azure resources

Pick names; storage / Key Vault / function names must be globally unique
(add a suffix if taken). Storage must be lowercase, <= 24 chars.

```powershell
$RG="rg-cat-broker"; $LOC="eastus"
$STORAGE="catbrokerstg01"; $KV="cat-broker-kv01"; $FUNC="cat-broker-func01"

az group create -n $RG -l $LOC
az storage account create -n $STORAGE -g $RG -l $LOC --sku Standard_LRS
az keyvault create -n $KV -g $RG -l $LOC

az functionapp create -n $FUNC -g $RG `
  --storage-account $STORAGE `
  --consumption-plan-location $LOC `
  --runtime python --runtime-version 3.11 `
  --functions-version 4 --os-type Linux
```

---

## Phase 3 — Wire the secret through Key Vault

```powershell
# managed identity for the function
az functionapp identity assign -n $FUNC -g $RG
$PRINCIPAL = az functionapp identity show -n $FUNC -g $RG --query principalId -o tsv
az keyvault set-policy -n $KV --object-id $PRINCIPAL --secret-permissions get list

# store the Cat credentials in Key Vault
az keyvault secret set --vault-name $KV --name CatClientId     --value "<your-client-id>"
az keyvault secret set --vault-name $KV --name CatClientSecret --value "<your-client-secret>"

# app settings: Key Vault references for the secrets, plain values for the rest
az functionapp config appsettings set -n $FUNC -g $RG --settings `
  "CAT_CLIENT_ID=@Microsoft.KeyVault(SecretUri=https://$KV.vault.azure.net/secrets/CatClientId/)" `
  "CAT_CLIENT_SECRET=@Microsoft.KeyVault(SecretUri=https://$KV.vault.azure.net/secrets/CatClientSecret/)" `
  "CAT_SCOPE=<your-client-id>/.default" `
  "CAT_TENANT_ID=ceb177bf-013b-49ab-8a9c-4abce32afc1e" `
  "CAT_DEFAULT_PARTY_NUMBER=ZZIO"
```

---

## Phase 4 — Deploy and test

```powershell
func azure functionapp publish $FUNC

# get the function key and call it
$KEY = az functionapp keys list -n $FUNC -g $RG --query "functionKeys.default" -o tsv
curl "https://$FUNC.azurewebsites.net/api/search?serial=9303&code=$KEY"
```

---

## Phase 5 (later) — per-user auth + clients

- **Entra Easy Auth:** Function App → Settings → Authentication → Add identity
  provider → Microsoft → restrict to your tenant (and a group if you like). This
  upgrades from a shared function key to per-user sign-in.
- **VBA workbook:** point `CatSearch` at `https://<func>.azurewebsites.net/api/search`
  with the function key, and remove the Cat secret from the Config sheet.
- **Office add-in:** use Office SSO to call the broker with the user's Entra token.

---

## Cost

Consumption plan: thousands of searches/month sit inside the free grant —
effectively **$0-$5/month** (mostly the storage account). Only move to a warm
plan (Basic App Service ~$13-55, or Premium ~$150) if cold-start latency bothers
users. Do **not** add API Management unless you need an enterprise gateway.
