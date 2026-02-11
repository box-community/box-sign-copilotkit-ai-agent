# Box Sign AI Assistant with CopilotKit

A conversational AI demo that transforms Box Sign e-signature workflows into natural language interactions. Built with **Box Node SDK v10**, **Box Content Preview**, **CopilotKit**, and Next.js.

> **Note**: This is a proof-of-concept demo. Before deploying to production, adapt it to your enterprise needs, implement proper authentication, add comprehensive error handling, and test thoroughly.

## Features

- **Natural language interface**: Create complex signature workflows by describing what you want in plain English
- **Document preview**: See the document and request details before sending (Box Content Preview integration)
- **Human-in-the-loop**: Explicit confirmation required before sending signature requests
- **Multiple participant roles**: Support for signers, approvers, and final copy readers
- **Advanced workflows**: Sequential signing, custom expiration, automatic reminders, custom email subjects/messages
- **Bulk operations**: Cancel multiple requests with a single command
- **Real-time tracking**: Auto-refreshing panel shows all active signature requests
- **Context persistence**: Multi-turn conversations with maintained context

## How It Works

- **CopilotKit**: Frontend actions (`useCopilotAction`) define tools that the LLM can invoke based on user intent
- **Box Sign API**: Server-side integration via [Box Node SDK v10](https://www.npmjs.com/package/box-node-sdk) for signature request management
- **Box Content Preview**: Displays document preview with request details for user confirmation before sending
- **Action orchestration**: LLM decides when to search files, prepare requests, create signatures, or check status based on conversation

## Prerequisites

- Node.js 18+
- [Box Developer Account](https://developer.box.com/) with a Custom App (Server Auth or Developer Token)
- Box app with **Read/Write files** and **Manage signature requests** (Box Sign) access
- **CORS configured** in your Box app: Add `http://localhost:3000` (no trailing slash) to CORS Domains in Developer Console
- Access to Box Sign API (available on Business plans and above)
- [OpenAI API key](https://platform.openai.com/) for CopilotKit

## Quick Start

1. **Install**

   ```bash
   npm install
   ```

2. **Environment**

   Copy `.env.example` to `.env` and set:

   - `BOX_DEVELOPER_TOKEN` – from Box Developer Console → Your App → Configuration (Developer Token)
   - `OPENAI_API_KEY` – for CopilotKit

3. **Run**

   ```bash
   npm run dev
   ```

   Open [http://localhost:3000](http://localhost:3000) and start chatting with the assistant.

## Usage Examples

### Simple signature request

```
Find my employment contract and send it to alice@company.com to sign
```

The assistant will search Box, show you the document preview, and ask for confirmation before sending.

### Advanced workflow with multiple roles

```
Search for Vendor agreement.pdf, then create a signature request:
- Approver: legal@company.com
- Signer: finance@company.com
Make it sequential, valid for 30 days, enable reminders, and set the subject to "Q4 Vendor Agreement – please sign"
```

This demonstrates:
- **Search**: Find documents by name
- **Multiple roles**: Approver (reviews/approves) and Signer (signs document)
- **Sequential order**: Legal approves first, then finance signs
- **Expiration**: 30-day validity
- **Reminders**: Automatic reminder emails
- **Custom email**: Personalized subject line

### Bulk operations

```
Cancel the last 10 signature requests for "vendor agreement"
```

Cancel multiple requests matching a search phrase.

### Status tracking

```
What's the status of my signature requests?
```

Or check a specific request:

```
Show me details of the request I just created
```

### Different participant roles

Box Sign supports three roles:

- **Signer**: Must sign the document
- **Approver**: Reviews and approves/declines (doesn't sign)
- **Final copy reader**: Only receives a copy when complete (no action required)

Example with all three:

```
Send the policy document to legal@co.com for approval, then manager@co.com to sign, and hr@co.com should get a final copy
```

> **Note**: Replace file names and email addresses with ones that exist in your Box account.

## Project Structure

```
├── app/
│   ├── api/
│   │   ├── copilotkit/route.ts   # CopilotKit runtime (OpenAI)
│   │   └── box/
│   │       └── route.ts          # Box API route: search, create/list/get/cancel/resend
│   ├── layout.tsx
│   ├── page.tsx
│   └── globals.css
├── components/
│   ├── BoxSignAssistant.tsx      # Main component with CopilotKit actions
│   └── BoxPreviewPanel.tsx       # Document preview with Box Content Preview
├── lib/
│   └── box-client.ts             # Box Node SDK v10 client (Developer Token)
├── .env.local.example
├── .gitignore
└── package.json
```

## CopilotKit Actions

The assistant exposes the following actions to the LLM:

- `search_files` - Find documents in Box by name
- `prepare_signature_request` - Show document preview and request details (requires confirmation)
- `create_signature_request` - Create Box Sign request with all parameters
- `list_signature_requests` - List all signature requests
- `get_signature_request_status` - Get details for a specific request
- `get_latest_signature_request_status` - Get details for the most recent request
- `cancel_signature_request` - Cancel a pending request
- `bulk_cancel_signature_requests` - Cancel multiple requests matching a search phrase
- `resend_signature_request` - Resend reminder emails

## Box Node SDK v10

This project uses **box-node-sdk v10**:

- **Auth**: `BoxDeveloperTokenAuth` + `BoxClient` (see `lib/box-client.ts`)
- **Sign**: `client.signRequests.createSignRequest(requestBody)` with `FileBase`, `FolderMini`, and `SignRequestCreateSigner` from `box-node-sdk/schemas`
- **Search**: `client.search.searchForContent({ query, type: 'file', limit })`
- **Sign requests**: `getSignRequests`, `getSignRequestById`, `cancelSignRequest`, `resendSignRequest`
- **File info**: `client.files.getFileById()` for preview tokens

All Box calls run server-side in API routes; the frontend only calls those routes from CopilotKit actions.

## Security Notes

- Use **Developer Token** only for local/testing. For production, use OAuth 2.0 or another supported auth method from the [Box Node SDK v10 docs](https://github.com/box/box-node-sdk)
- Never commit `.env.local` or real tokens (already excluded in `.gitignore`)
- Box Sign requires [Sign API access](https://developer.box.com/guides/box-sign/) to be enabled for your app
- Document content is sent to OpenAI for LLM processing - evaluate this for your use case
- Consider implementing audit logging for signature requests in production
- Add proper error handling and validation before production use

## Troubleshooting

### 403 "insufficient_scope" from Box

If you see `403 "insufficient_scope" "The request requires higher privileges than provided by the access token."`, your **Box app's application scopes** are missing something this demo needs.

1. Open [Box Developer Console](https://app.box.com/developers/console) → your app → **Configuration**
2. Under **Application Scopes**, enable:
   - **Read all files and folders stored in Box**
   - **Write all files and folders stored in Box**
   - **Manage signature requests** (required for Box Sign)
3. If your app has **Advanced Features**, ensure **Box Sign** (or "Manage signature requests") is enabled there as well
4. **Generate a new Developer Token** after changing scopes (existing tokens keep the old scopes). Copy the new token into `BOX_DEVELOPER_TOKEN` in `.env.local` and restart the app

See [Box Sign – Box Dev Docs](https://developer.box.com/guides/box-sign) and [Scopes](https://box.dev/guides/api-calls/permissions-and-errors/scopes) for details.

### CORS Issues with Box Content Preview

If you see CORS errors when previewing documents:

1. Open [Box Developer Console](https://app.box.com/developers/console) → your app → **Configuration**
2. Under **CORS Domains**, add `http://localhost:3000`
3. Save changes and restart your dev server

### Missing Box Sign Access

If Box Sign features are not available:

- Box Sign is available on **Business plans and above**
- Verify your account has access at [Box Admin Console](https://app.box.com/master/settings)
- Contact your Box administrator if needed

## Resources

- [Box Sign API](https://developer.box.com/reference/post-sign-requests)
- [Box Node SDK v10](https://www.npmjs.com/package/box-node-sdk)
- [Box Content Preview](https://developer.box.com/guides/embed/ui-elements/preview/)
- [CopilotKit Documentation](https://docs.copilotkit.ai/)
- [Box Developer Community](https://community.box.com/)

## License

MIT
