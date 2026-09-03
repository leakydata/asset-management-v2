# Cat Asset Management — broker (Azure Function)

> ## ⚠ THIS SOURCE IS BEHIND WHAT IS DEPLOYED
>
> `function_app.py` here defines **one route: `search`**. The Excel add-in also
> calls **`ownership`**, **`expire`** and **`transfer`**, and they work — so the
> deployed Function has at least three routes that are not in this folder.
>
> **Do not treat this folder as the source of truth**, and do not redeploy from
> it: doing so would remove the three write routes and break every operation
> sheet. Get the current source from the deployed Function App
> (`asset-management-proxy-…` in Azure) before changing anything.
>
> Verified 2026-09-03 by calling the live endpoint.

A small HTTP Function that holds the Cat API credentials server-side, so the
client secret is never distributed. Excel calls this instead of Cat.

`function_app.py` reuses the same OAuth client-credentials + search logic as the
MCP server.

## What Cat's search can and cannot filter on

Worth knowing before designing anything that needs to *find* records, because it
is the constraint that shapes what is possible:

From `catDigitalPlatform-assetManagement-v2-oas_prod.yaml`, `/ownershipRecords/search`:

- filters are `stringEquals` only, **maximum two of them**
- searchable on **DCN, asset name, serial number, make code** — and make code
  alone is not enough, it needs a second filter
- **no filter on `ownershipRequestType`, dealer code, or record status**

So there is **no way to ask "everything pending for our dealer code"** — not
from the add-in, and not from a new proxy route either. The API simply does not
offer that query.

What *does* work: search returns **ACTIVE and PENDING records both**, and
`ownershipRequestType` is populated on the pending ones. So "what is waiting on
me" is answerable by sweeping a known list of DCNs or serials and filtering the
results client-side — no proxy change needed.

Confirmed against the live endpoint: `?status=PENDING` alone returns 400, and
`?dcn=…&status=PENDING` returns exactly the same records as `?dcn=…` — unknown
parameters are silently ignored rather than rejected.

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
