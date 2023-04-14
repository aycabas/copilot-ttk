const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Message extension Code
  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {
    const searchQuery = query.parameters[0].value;
    // const response = await axios.get(
    //   `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
    //     text: searchQuery,
    //     size: 8,
    //   })}`
    // );
    //get response from the api call to get products by name using port 3978
    const response = await axios.get(
      `http://localhost:3978/api/products/${searchQuery}`
    );

    const attachments = [];
    response.data.forEach((obj) => {
      const heroCard = CardFactory.heroCard(obj.name);
      const preview = CardFactory.heroCard(obj.name);
      preview.content.tap = {
        type: "invoke",
        value: { name: obj.name, description: obj.description },
      };
      const attachment = { ...heroCard, preview };
      attachments.push(attachment);
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }

  async handleTeamsMessagingExtensionSelectItem(context, obj) {
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [CardFactory.heroCard(obj.name, obj.description)],
      },
    };
  }
}

module.exports.TeamsBot = TeamsBot;
