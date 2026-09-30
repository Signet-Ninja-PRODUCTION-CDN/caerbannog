export default {
    
    async fetch(request, env) {
        
        // ============================================================
        // INVOCATION METADATA
        // ============================================================
        
        const invocationGUID = crypto.randomUUID();
        
        const submitTDS = new Date().toISOString();
        
        const rayID = request.headers.get("CF-Ray") || "";
        
        const originURL = request.headers.get("Referer") || request.headers.get("Origin") || "";
        
        const geoRegion = request.cf?.region || "";
        
        const geoCity = request.cf?.city || "";
        
        const geoCountry = request.cf?.country || "";
        
        const deviceAgent = request.headers.get("User-Agent") || "";
        
        // ============================================================
        // CORS PREFLIGHT
        // ============================================================
        
        if (request.method === "OPTIONS") {
            return new Response(null, {
                status: 204,
                headers: corsHeaders()
            });
        }
        
        const startedAt = Date.now();
        
        console.log(JSON.stringify({
            event: "request.accepted",
            requestID: invocationGUID,
            rayID: rayID,
            method: request.method
        }));

        // ============================================================
        // POST ONLY
        // ============================================================

        if (request.method !== "POST") {
            console.warn(JSON.stringify({
                event: "request.rejected",
                requestID: invocationGUID,
                rayID: rayID,
                reason: "method_not_allowed",
                method: request.method,
                httpStatus: 405,
                duration_ms: Date.now() - startedAt
            }));
        
            return jsonResponse(
                {
                    success: false,
                    requestID: invocationGUID,
                    error: "Method not allowed"
                },
                405
            );
        }

        // ============================================================
        // REQUIRE JSON
        // ============================================================

        const contentType = request.headers.get("Content-Type") || "";
    
        if (!contentType.toLowerCase().startsWith("application/json")) {
            console.warn(JSON.stringify({
                event: "request.rejected",
                requestID: invocationGUID,
                rayID: rayID,
                reason: "unsupported_media_type",
                httpStatus: 415,
                duration_ms: Date.now() - startedAt
            }));
    
            return jsonResponse(
                {
                    success: false,
                    requestID: invocationGUID,
                    error: "Unsupported media type"
                },
                415
            );
        }

        // ============================================================
        // PARSE SUBMISSION
        // ============================================================

        let payload;
    
        try {
            payload = await request.json();
        }
        
        catch {

            console.warn(JSON.stringify({
                event: "request.rejected",
                requestID: invocationGUID,
                rayID: rayID,
                reason: "invalid_json",
                httpStatus: 400,
                duration_ms: Date.now() - startedAt
            }));
            
            return jsonResponse(
                {
                    success: false,
                    requestID: invocationGUID,
                    error: "Invalid JSON"
                },
                400
            );
        }
        
        // ============================================================
        // MULTICONNECT FORM FIELDS
        // ============================================================

        const formName = String(payload["formmsg-submit.name"] || "").trim();

        const formEmail = String(payload["formmsg-submit.email"] || "").trim();

        const formSubj = String(payload["formmsg-submit.subj"] || "").trim();

        const formBody = String(payload["formmsg-submit.body"] || "").trim();

        const formListSU = payload["formmsg-submit.sub.genann"] === true;

        // ============================================================
        // BASIC FORM VALIDATION
        // ============================================================

        if (!formName || !formEmail || !formSubj || !formBody) {
            console.warn(JSON.stringify({
                event: "request.rejected",
                requestID: invocationGUID,
                rayID: rayID,
                reason: "missing_required_field",
                httpStatus: 400,
                duration_ms: Date.now() - startedAt
            }));
        
            return jsonResponse(
                {
                    success: false,
                    requestID: invocationGUID,
                    error: "Missing required form field"
                },
                400
            );
        }

        // ============================================================
        // MICROSOFT GRAPH ACCESS TOKEN
        // ============================================================

        let tokenResponse;
    
        try {
            tokenResponse = await fetch(
                `https://login.microsoftonline.com/${env.EntraTenID}/oauth2/v2.0/token`,
                {
                    method: "POST",
                    
                    headers: {
                        "Content-Type": "application/x-www-form-urlencoded"
                    },
                
                    body: new URLSearchParams({
                        client_id: env.EntraCliID,
                        client_secret: env.EntraCliSec,
                        scope: "https://graph.microsoft.com/.default",
                        grant_type: "client_credentials"
                    })
                }
            );
        }
    
        catch (error) {
            console.error(JSON.stringify({
                event: "graph.token.transport_failure",
                requestID: invocationGUID,
                rayID: rayID,
                outcome: "failure",
                error: String(error),
                httpStatus: 502,
                duration_ms: Date.now() - startedAt
            }));
        
            return ingestFailure(invocationGUID);
        }
    
        if (!tokenResponse.ok) {
            console.error(JSON.stringify({
                event: "graph.token.rejected",
                requestID: invocationGUID,
                rayID: rayID,
                outcome: "failure",
                upstreamStatus: tokenResponse.status,
                httpStatus: 502,
                duration_ms: Date.now() - startedAt
            }));
        
            return ingestFailure(invocationGUID);
        }
    
        const tokenResult = await tokenResponse.json();

        const accessToken = tokenResult.access_token;
    
        console.log(JSON.stringify({
            event: "graph.token.success",
            requestID: invocationGUID,
            rayID: rayID,
            outcome: "success",
            duration_ms: Date.now() - startedAt
        }));

        // ============================================================
        // SHAREPOINT RECORD
        // These are the Graph-confirmed SharePoint field names.
        // ============================================================

        const graphBody = {
            fields: {
                Title:
                "MultiConnect General Message Submission",
            
                ingest_cfworkerInvokeGUID:
                invocationGUID,
            
                ingest_cfworkerRayID:
                rayID,
            
                ingest_cfworkerSubmitTDS:
                submitTDS,
            
                ingest_cfworkerOriginURL:
                originURL,
            
                ingest_cfworkerGeoRegion:
                geoRegion,
            
                ingest_cfworkerGeoCity:
                geoCity,
            
                ingest_cfworkerGeoCountry:
                geoCountry,
            
                ingest_cfworkerDeviceAgent:
                deviceAgent,
            
                ingest_FormEmail:
                formEmail,
            
                ingest_FormName:
                formName,
            
                ingest_FormSubj:
                formSubj,
            
                ingest_FormBody:
                formBody,
            
                ingest_FormListSU:
                formListSU
            
                // Reserved for Future Utilization
                // // ingest_SPRecordGUID:
                // // ingest_SPRecordStat:
            
            }
        };

        // ============================================================
        // CREATE SHAREPOINT LIST ITEM
        // ============================================================

        const graphURL = `https://graph.microsoft.com/v1.0/sites/${env.SPGraphSiteID}/lists/${env.SPGraphListID}/items`;
    
        let graphResponse;
    
        try {
            graphResponse = await fetch(
                graphURL,
                {
                    method: "POST",
                    
                    headers: {
                        "Authorization": `Bearer ${accessToken}`,
                        "Content-Type": "application/json"
                    },
                    
                    body: JSON.stringify(graphBody)
                }
            );
        }
    
        catch (error) {
            console.error(JSON.stringify({
                event: "sharepoint.create.transport_failure",
                requestID: invocationGUID,
                rayID: rayID,
                outcome: "failure",
                error: String(error),
                httpStatus: 502,
                duration_ms: Date.now() - startedAt
            }));
            
            return ingestFailure(invocationGUID);
        }
    
        const graphResponseText = await graphResponse.text();
    
        if (!graphResponse.ok) {
            console.error(JSON.stringify({
                event: "sharepoint.create.rejected",
                requestID: invocationGUID,
                rayID: rayID,
                outcome: "failure",
                upstreamStatus: graphResponse.status,
                httpStatus: 502,
                duration_ms: Date.now() - startedAt
            }));
        
            return ingestFailure(invocationGUID);
        }

        // ============================================================
        // SUCCESS LOGGING
        // ============================================================

        let sharePointItemID = null;
    
        try {
            const graphResult =
            JSON.parse(graphResponseText);
        
            sharePointItemID =
            graphResult.id || null;
        }
    
        catch {
            // Graph already returned success.
            // Failure to parse its response does not invalidate the ingest.
        }
        
        console.log(JSON.stringify({
            event: "sharepoint.create.success",
            requestID: invocationGUID,
            rayID: rayID,
            sharePointItemID: sharePointItemID,
            graphStatus: graphResponse.status
        }));
        
        console.log(JSON.stringify({
            event: "request.completed",
            requestID: invocationGUID,
            rayID: rayID,
            outcome: "success",
            httpStatus: 201,
            duration_ms: Date.now() - startedAt
        }));

        // ============================================================
        // CLIENT RESPONSE
        // ============================================================

        return jsonResponse(
            {
                success: true,
                requestID: invocationGUID
            },
            201
        );
    }
};

// ================================================================
// RESPONSE HELPERS
// ================================================================

function corsHeaders() {
    return {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type"
    };
}

function jsonResponse(body, status) {
    return new Response(
        JSON.stringify(body),
        {
            status,
            headers: {
                "Content-Type": "application/json; charset=utf-8",
                ...corsHeaders()
            }
        }
    );
}

function ingestFailure(invocationGUID) {
    return jsonResponse(
        {
            success: false,
            requestID: invocationGUID,
            error: "Message ingest failed"
        },
        502
    );
}