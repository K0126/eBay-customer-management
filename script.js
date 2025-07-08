function updatePaidAndShippedOrders() {
  Logger.log("스크립트 시작");

  const token = "ENTER YOUR EBAY TOKEN"; //2026년에 만료, 발급은 ebay developer portal에서 
  const url = "https://api.ebay.com/ws/api.dll"; // limit 25000/day
  const now = new Date();
  const createTimeTo = now.toISOString();
  const createTimeFrom = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000).toISOString(); // 최근 7일간 내역만 파싱



  Logger.log("eBay API 요청 준비 중");


  const xmlRequest = `
    <?xml version="1.0" encoding="utf-8"?>
    <GetOrdersRequest xmlns="urn:ebay:apis:eBLBaseComponents">
      <RequesterCredentials>
        <eBayAuthToken>${token}</eBayAuthToken>
      </RequesterCredentials>
      <CreateTimeFrom>${createTimeFrom}</CreateTimeFrom>
      <CreateTimeTo>${createTimeTo}</CreateTimeTo>
      <OrderRole>Seller</OrderRole>
      <OrderStatus>Completed</OrderStatus>
    </GetOrdersRequest>`;

  const options = {
    method: "post",
    contentType: "text/xml",
    payload: xmlRequest,
    headers: {
      "X-EBAY-API-CALL-NAME": "GetOrders",
      "X-EBAY-API-SITEID": "0",
      "X-EBAY-API-COMPATIBILITY-LEVEL": "967"
    },
    muteHttpExceptions: true
  };

  let response, xmlText, document;
  try {
    response = UrlFetchApp.fetch(url, options);
    xmlText = response.getContentText();
    document = XmlService.parse(xmlText);
    Logger.log("eBay 응답 성공");
  } catch (e) {
    Logger.log("eBay 요청 또는 파싱 실패: " + e);
    return;
  }

  const root = document.getRootElement();
  const ns = XmlService.getNamespace("urn:ebay:apis:eBLBaseComponents");
  const orderArray = root.getChild("OrderArray", ns);
  if (!orderArray) {
    Logger.log("OrderArray 없음");
    return;
  }

  const orders = orderArray.getChildren("Order", ns);
  Logger.log(`총 주문 수: ${orders.length}`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ebay (Margaret)"); // 시트 이름 수정시 수정
  if (!sheet) {
    Logger.log("시트를 찾을 수 없음");
    return;
  }



  const columnD = sheet.getRange("D:D").getValues().map(r => r[0]).filter(v => v); // 마지막 주문은 D/E 행에서만 찾게 되어있음, 추후 다른 행으로 넘어갈 시 이걸 수정해줘야함
  const columnE = sheet.getRange("E:E").getValues().map(r => r[0]).filter(v => v);
  const lastRow = columnD.length;
  const lastName = columnD[lastRow - 1]?.toString().trim().toLowerCase();
  const lastEmail = columnE[lastRow - 1]?.toString().trim().toLowerCase();

  Logger.log(`시트 마지막 고객: ${lastName} / ${lastEmail}`);

  const existingSet = new Set();
  for (let i = 0; i < columnD.length; i++) {
    const name = columnD[i]?.toString().trim().toLowerCase();
    const email = columnE[i]?.toString().trim().toLowerCase();
    if (name && email) existingSet.add(`${name}|${email}`);
  }

  let foundIndex = -1;
  for (let i = 0; i < orders.length; i++) {
    const order = orders[i];

    // Get buyer name
    let buyerName = "";
    const transactions = order.getChild("TransactionArray", ns);
    if (transactions) {
      const transactionList = transactions.getChildren("Transaction", ns);
      if (transactionList.length > 0) {
        const buyer = transactionList[0].getChild("Buyer", ns);
        buyerName = buyer?.getChildText("UserFirstName", ns)?.trim() || "";
        const lastNameBuyer = buyer?.getChildText("UserLastName", ns)?.trim() || "";
        if (buyerName && lastNameBuyer) {
          buyerName += ` ${lastNameBuyer}`;
        } else if (lastNameBuyer) {
          buyerName = lastNameBuyer;
        }
      }
    }
    const name = buyerName.toLowerCase();

    let email = "";
    if (transactions) {
      const transactionList = transactions.getChildren("Transaction", ns);
      if (transactionList.length > 0) {
        const buyer = transactionList[0].getChild("Buyer", ns);
        email = buyer?.getChildText("Email", ns)?.trim().toLowerCase() || "";
      }
    }

    if (name === lastName && email === lastEmail) {
      foundIndex = i;
      Logger.log(`마지막 고객 인덱스: ${i}`);
      break;
    }
  }

  if (foundIndex === -1) {
    Logger.log("API 응답에서 기준 고객을 찾을 수 없음 — 추가 없음"); // 오류
    return;
  }

  let added = 0;
  let skipped = 0;

  for (let i = foundIndex + 1; i < orders.length; i++) {
    const order = orders[i];
    const paidTime = order.getChildText("PaidTime", ns);
    const shippedTime = order.getChildText("ShippedTime", ns);
    if (!paidTime || !shippedTime) continue;

    // Get buyer name for new entries
    let buyerName = "N/A";
    const transactions = order.getChild("TransactionArray", ns);
    if (transactions) {
      const list = transactions.getChildren("Transaction", ns);
      if (list.length > 0) {
        const buyer = list[0].getChild("Buyer", ns);
        const firstName = buyer?.getChildText("UserFirstName", ns)?.trim() || "";
        const lastName = buyer?.getChildText("UserLastName", ns)?.trim() || "";
        if (firstName && lastName) {
          buyerName = `${firstName} ${lastName}`;
        } else if (firstName) {
          buyerName = firstName;
        } else if (lastName) {
          buyerName = lastName;
        }
      }
    }
    const name = buyerName;


    let email = "N/A";
    if (transactions) {
      const list = transactions.getChildren("Transaction", ns);
      if (list.length > 0) {
        const buyer = list[0].getChild("Buyer", ns);
        email = buyer?.getChildText("Email", ns)?.trim() || "N/A";
      }
    }

    const key = `${name.toLowerCase()}|${email.toLowerCase()}`;
    if (existingSet.has(key)) {
      skipped++;
      continue;
    }

    sheet.getRange(lastRow + added + 1, 4).setValue(name);
    sheet.getRange(lastRow + added + 1, 5).setValue(email);
    existingSet.add(key);
    added++;
  }

  Logger.log(`✅ 완료: ${added}명 추가됨 / 🔁 ${skipped}명 중복 스킵됨`);
}
