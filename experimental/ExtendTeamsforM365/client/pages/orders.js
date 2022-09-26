import { getOrder } from "../modules/northwindDataService.js";
import chatCard from "../cards/orderChatCard.js";
import orderTrackerCard from "../cards/orderTrackerCard.js";
import mailCard from "../cards/orderMailCard.js";
import { env } from "/modules/env.js";
import { inM365 } from "../modules/teamsHelpers.js";
//import { getOrders } from "../../server/northwindDataService.js";
async function displayUI() {
  let orderDetails = {};
  const displayElement = document.getElementById("content");
  const detailsElement = document.getElementById("orderDetails");
  try {
    const searchParams = new URLSearchParams(window.location.search);
    if (searchParams.has("orderId")) {
      const orderId = searchParams.get("orderId");
      const order = await getOrder(orderId);
      orderDetails.orderId = orderId ? orderId : "";
      orderDetails.contact =
        order.contactName && order.contactTitle
          ? `${order.contactName}(${order.contactTitle})`
          : "";
      //get from graph, for use env config with other users in your AAD
      orderDetails.salesRepEmail =
        env.CONTACTS.indexOf(",") > -1
          ? env.CONTACTS.split(",")
          : [env.CONTACTS];
      orderDetails.salesRepMailrecipients = env.CONTACTS.replace(",", ";");
      displayElement.innerHTML = `
                    <h2>Order details for ${order.orderId}</h2>
                    <p><b>Customer:</b> ${order.customerName}<br />
                    <b>Contact:</b> ${order.contactName}, ${
        order.contactTitle
      }<br />
                    <b>Date:</b> ${new Date(order.orderDate).toLocaleDateString(
                      "en-us",
                      {
                        weekday: "short",
                        year: "numeric",
                        month: "short",
                        day: "numeric",
                      }
                    )}<br />
                    <b> ${order.employeeTitle}</b>: ${order.employeeName} (${
        order.employeeId
      })
                    </p>
                `;
      order.details.forEach((item) => {
        const orderRow = document.createElement("tr");
        orderRow.innerHTML = `<tr>
                        <td>${item.quantity}</td>
                        <td><a href="/pages/productDetail.html?productId=${item.productId}">${item.productName}</a></td>
                        <td>${item.unitPrice}</td>
                        <td>${item.discount}</td>
                    </tr>`;
        detailsElement.append(orderRow);
      });
      //show tracker element for each order
      const trackerArea = document.getElementById("trackerBox");
      trackerArea.style.display = "block";
      var template = new ACData.Template(orderTrackerCard);
      var card = template.expand({ $root: orderDetails });
      var adaptiveCard = new AdaptiveCards.AdaptiveCard();
      adaptiveCard.parse(card);
      trackerArea.appendChild(adaptiveCard.render());
      if (await inM365()) {
        //chat support
        if (microsoftTeams.chat.isSupported()) {
          //show chat view
          const chatArea = document.getElementById("chatBox");
          chatArea.style.display = "block";

          //adaptive card templating
          var template = new ACData.Template(chatCard);
          var card = template.expand({ $root: orderDetails });
          var adaptiveCard = new AdaptiveCards.AdaptiveCard();

          //button action for chat
          adaptiveCard.onExecuteAction = async (action) => {
            if (orderDetails.salesRepEmail.length > 1) {
              //group chat
              await microsoftTeams.chat.openGroupChat({
                users: orderDetails.salesRepEmail,
                topic: `Enquiry about order ${orderDetails.orderId}`,
                message: `Hi, to discuss about ${orderDetails.orderId}`,
              });
            } else {
              //1:1 chat
              await microsoftTeams.chat.openChat({
                user: orderDetails.salesRepEmail[0],
                message: `Enquiry about order ${orderDetails.orderId}`,
              });
            }
          };
          adaptiveCard.parse(card);
          chatArea.appendChild(adaptiveCard.render());
        } else if (microsoftTeams.mail.isSupported()) {
          //show mail view
          const mailArea = document.getElementById("mailBox");
          mailArea.style.display = "block";

          //adaptive card templating
          var template = new ACData.Template(mailCard);
          var card = template.expand({ $root: orderDetails });
          var adaptiveCard = new AdaptiveCards.AdaptiveCard();

          //button action for new mail
          adaptiveCard.onExecuteAction = (action) => {
            microsoftTeams.mail.composeMail({
              type: microsoftTeams.mail.ComposeMailType.New,
              subject: `Enquire about order ${orderDetails.orderId}`,
              toRecipients: [orderDetails.salesRepMailrecipients],
              message: "Hello",
            });
          };

          adaptiveCard.parse(card);
          mailArea.appendChild(adaptiveCard.render());
        } else {
          message.innerText = `Error: chat/mail not supported`;
        }
      }
    } else {
      const orders = [];
      orders.push(await getOrder(10248));
      orders.push(await getOrder(10249));
      orders.push(await getOrder(10250));
      orders.push(await getOrder(10251));
      orders.push(await getOrder(10252));
      orders.push(await getOrder(10253));
      orders.push(await getOrder(10254));
      orders.push(await getOrder(10255));
      orders.push(await getOrder(10256));
      orders.push(await getOrder(10257));
      orders.push(await getOrder(10258));
      orders.push(await getOrder(10259));
      orders.push(await getOrder(10260));
      displayElement.innerHTML = `
                <h2>Showing order details</h2>
            `;

      orders.forEach((orderRecord) => {
        const orderRow = document.createElement("tr");
        const pageRef = `orderDetail.html?orderId=${orderRecord.orderId}`;
        orderRow.innerHTML = `<td><button onClick="location.href='${pageRef}'">${orderRecord.orderId}</button></td>
                                    <td>${orderRecord.customerName}</td>
                                    <td>${orderRecord.details[0].productName}</td>
                                    <td>${orderRecord.details[0].quantity}</td>
                                    <td>${orderRecord.details[0].categoryName}</td>
                                    <td>'Shipped'</td>`;
        detailsElement.append(orderRow);
      });
    }
  } catch (error) {
    // If here, we had some other error
    message.innerText = `Error: ${JSON.stringify(error)}`;
  }
}

//display the tab for order details
displayUI();
