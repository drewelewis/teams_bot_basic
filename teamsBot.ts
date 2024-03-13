import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
} from "botbuilder";
import config from "./config";

export class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      context.sendActivity({type: "typing"})
      const aiResponse = await getAIResponse(txt)

      await context.sendActivity(`${aiResponse}`);
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity(
            `Hi there! How can I help you?`
          );
          break;
        }
      }
      await next();
    });
  }
}
interface Response {
  text: string
}
function getAIResponse(query: string): Promise<Response[]> {

  const url=config.apiwrapper+query
  // We can use the `Headers` constructor to create headers
  // and assign it as the type of the `headers` variable
  const headers: Headers = new Headers()
  // Add a few headers
  headers.set('Content-Type', 'application/json')
  headers.set('Accept', 'application/json')

  // Create the request object, which will be a RequestInfo type. 
  // Here, we will pass in the URL as well as the options object as parameters.
  const request: RequestInfo = new Request(url, {
    method: 'GET',
    headers: headers
  })

  // Pass in the request object to the `fetch` API
  return fetch(request)
    .then(res => res.json())
    .then(res => {
      return res as Response[]
    })
}
