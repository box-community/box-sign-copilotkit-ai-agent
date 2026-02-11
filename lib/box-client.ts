/**
 * Box API client using Box Node SDK v10.
 * Uses Developer Token auth (for demo). Use OAuth 2.0 in production.
 */
import { BoxClient, BoxDeveloperTokenAuth } from "box-node-sdk";

function getBoxClient(): BoxClient {
  const token = process.env.BOX_DEVELOPER_TOKEN;
  if (!token) {
    throw new Error(
      "BOX_DEVELOPER_TOKEN is not set. Get a token from Box Developer Console."
    );
  }
  const auth = new BoxDeveloperTokenAuth({ token });
  return new BoxClient({ auth });
}

export { getBoxClient };
