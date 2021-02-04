const excelTOJSON = require("convert-excel-to-json");
const fs = require("fs");
const { resolve } = require("path");

const result = excelTOJSON({
  sourceFile: resolve(__dirname, "files", "sinapi.xls"),
  header: {
    rows: 8,
  },
  sheets: [
    {
      name: "Rel. Analítico",
      columnToKey: {
        G: "compositionCode",
        H: "compositionDescription",
        I: "und",
        M: "itemCode",
        N: "itemDescription",
        Q: "coef",
        R: "price",
      },
    },
  ],
});

const groupItems = (items) => {
  let newItems = [];

  items.map((item) => {
    const itemExistsInArray = newItems.findIndex(
      (i) => item.compositionCode === i.compositionCode
    );
    if (itemExistsInArray < 0) {
      if (item.itemDescription.length > 1) {
        newItems.push({
          compositionCode: item.compositionCode,
          compositionDescription: item.compositionDescription,
          items: [
            {
              itemCode: item.itemCode,
              itemDescription: item.itemDescription,
              und: item.und,
              coef: parseFloat(item.coef.replace(",", ".")),
              price: parseFloat(item.price.replace(",", ".")),
            },
          ],
        });
      }
    } else {
      if (item.itemDescription.length > 1) {
        newItems[itemExistsInArray].items.push({
          itemCode: item.itemCode,
          itemDescription: item.itemDescription,
          und: item.und,
          coef: parseFloat(item.coef.replace(",", ".")),
          price: parseFloat(item.price.replace(",", ".")),
        });
      }
    }
  });

  return newItems;
};

const orderedResult = groupItems(result["Rel. Analítico"]);

const resultInString = JSON.stringify(orderedResult);
fs.writeFileSync("sinapi-data.json", resultInString, (err, result) => {
  if (err) console.error(err);
  console.log("sucesso");
});
