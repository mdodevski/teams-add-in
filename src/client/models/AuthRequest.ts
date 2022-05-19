export interface AuthRequest{
    grant_type: string,
    client_id: string,
    client_secret: string,
    scope: string,
    requested_token_use: string,
    assertion: string
}