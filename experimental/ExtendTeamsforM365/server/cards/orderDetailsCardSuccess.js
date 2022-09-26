export default {
  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
  originator: "${originator}",
  type: "AdaptiveCard",
  version: "1.4",
  body: [
    {
      type: "ColumnSet",
      spacing: "small",
      columns: [
        {
          type: "Column",
          width: "stretch",
          items: [
            {
              type: "TextBlock",
              text: "Order Details - ${OrderId}",
              isSubtle: true,
              size: "medium",
              wrap: true,
              weight: "bolder",
              maxLines: 1,
            },
          ],
        },
      ],
    },
    {
      type: "ColumnSet",
      width: "stretch",
      columns: [
        {
          type: "Column",
          items: [
            {
              type: "FactSet",
              facts: [
                {
                  title: "Customer Name",
                  value: "${CustomerName}",
                },
                {
                  title: "Product Name",
                  value: "${ProductName}",
                },
                {
                  title: "Quantity",
                  value: "${Quantity}",
                },
                {
                  title: "Category",
                  value: "${Category}",
                },
                {
                  title: "Order Status",
                  value: "${OrderStatus}",
                },
              ],
            },
          ],
        },
      ],
    },
    {
      type: "ColumnSet",
      spacing: "small",
      columns: [
        {
          type: "Column",
          width: "stretch",
          items: [
            {
              type: "TextBlock",
              text: "Record updated successfully",
              isSubtle: true,
              size: "medium",
              wrap: true,
              maxLines: 1,
              color: "good",
            },
          ],
        },
      ],
    },
    {
      type: "ColumnSet",
      columns: [
        {
          type: "Column",
          items: [
            {
              type: "Input.Text",
              style: "text",
              id: "newQuantity",
              label: "New quantity count",
            },
          ],
        },
      ],
    },
  ],
  actions: [
    // {
    //   type: "Action.ShowCard",
    //   title: "Update Order",
    //   card: {
    //     type: "AdaptiveCard",
    //     body: [
    //       {
    //         type: "ColumnSet",
    //         columns: [
    //           {
    //             type: "Column",
    //             items: [
    //               {
    //                 type: "Input.Text",
    //                 style: "text",
    //                 id: "txtQuantity",
    //                 label: "New quantity count",
    //               },
    //             ],
    //           },
    //         ],
    //       },
    //     ],
    //     actions: [
    //       {
    //         type: "Action.Execute",
    //         title: "Update stock",
    //         verb: "OrderUpdate",
    //         data: {
    //           orderId: "${OrderId}",
    //         },
    //         style: "positive",
    //       },
    //     ],
    //   },
    // },
    {
      type: "Action.Execute",
      title: "Update stock",
      verb: "updateOrderDetails",
      data: {
        orderId: "${OrderId}",
        productId: "${ProductId}",
      },
      style: "positive",
    },
    {
      type: "Action.OpenUrl",
      title: "View order",
      url: "https://www.microsoft.com",
    },
  ],
};
