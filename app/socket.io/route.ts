import { NextRequest, NextResponse } from "next/server";

/**
 * This app does not use Socket.IO. Some clients (e.g. CopilotKit runtime or
 * devtools) may probe /socket.io. Return 404 so they get a clear response
 * instead of a 500 from an unhandled request.
 */
export async function GET(_request: NextRequest) {
  return new NextResponse("Not Found", { status: 404 });
}

export async function POST(_request: NextRequest) {
  return new NextResponse("Not Found", { status: 404 });
}
