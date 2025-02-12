function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('OpenAI_API_Key');
}

function GPT(prompt) {
  //'GPT'는 이 스크립트의 이름임. 나중에 호출할 때 사용하므로, 간단하게 설정.
  // 여러 개의 스크립트를 작성할 때는 'GPT1', 'GPT2' 등으로 구분하면 좋음.
  var apiKey = getApiKey();
  if (!apiKey) {
    return "Error: API key not set.";
  }
  if (!prompt) {
    return "Error: Please provide a valid prompt.";
  }
  
  // persistent cache: PropertiesService를 이용하여 이전 결과를 저장
  var properties = PropertiesService.getScriptProperties();
  // 프롬프트에 대해 고유한 키를 생성 (base64 인코딩 사용)
  var cacheKey = "gpt_" + Utilities.base64Encode(prompt);
  var cachedAnswer = properties.getProperty(cacheKey);
  if (cachedAnswer) {
    return cachedAnswer;  // 저장된 결과가 있으면 바로 반환
  }
  
  var url = "https://api.openai.com/v1/chat/completions";
  
  var payload = {
    "model": "gpt-4o-mini-2024-07-18",
    "messages": [
      {
        "role": "system",
        "content": "You are a bot that extracts main word keywords for product reviews." // Ai봇의 행동방침. 이 예시에서는 상품 리뷰의 키워드를 추출하는 역할이다.
      },
      {
        "role": "user",
        "content": prompt
      }
    ],
    "max_tokens": 200  // 가능한 출력 토큰 양. 50~500 사이 추천.
  };
  
  var options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + apiKey,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      return "Error: Request failed with status " + responseCode + ". " + response.getContentText();
    }
    var json = JSON.parse(response.getContentText());
    if (!json.choices || json.choices.length === 0) {
      return "Error: No response received from GPT.";
    }
    var answer = json.choices[0].message.content.trim();
    // 응답을 영구적으로 저장 (TTL 없이, 수동으로 삭제하지 않는 한 계속 유지됨)
    properties.setProperty(cacheKey, answer);
    return answer;
  } catch (e) {
    return "Error: " + e.message;
  }
}

function updateGPTResponses() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var question = sheet.getRange(i, 1).getValue();
    var answerCell = sheet.getRange(i, 2);
    if (question && answerCell.getValue() === "") {
      var answer = GPT(question);
      answerCell.setValue(answer);
      Utilities.sleep(1000);
    }
  } 
}


