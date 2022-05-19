import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/teamsAddInTab/index.html")
@PreventIframe("/teamsAddInTab/config.html")
@PreventIframe("/teamsAddInTab/remove.html")
export class TeamsAddInTab {
}
