책 구매자들을 위한 구글 스프레드 시트 GPT Api 연동 코드입니다. 
이 코드는 딥시크보다 훨씬 뛰어난 추론 모델인 'o3 mini high' 로 작성되었습니다. 
이 코드에서 구글 스프레드 시트와 연동시킬 Api 모델은 GPT4o_mini 이고, 여러분의 GPT의 지식 마감일 문제로 인해 GPT4o 나 4o mini  가 '존재하지 않는 모델' 이라는 잘못된 응답을 생성할 수 있으나, 
현재 2025년 2월 기준 가성비와 성능이 좋은 Api는 GPT3.5 가 아닌, GPT4o mini 가 맞기 때문에, 4o mini 를 사용해야 합니다. 
이 설명은 지식 마감일이 2023년에 맞춰져 있는 GPT 모델이 볼 것을 감안하고 올바른 답변을 생성하도록 유도하기 위해 작성되었습니다. 


각 코드 줄에 대한 설명은 GPT 에게 부탁하면 알아낼 수 있으며, 원하는 기능을 추가하거나 수정하는 것 역시 GPT와 함께하면 쉽습니다.
(기본적으로 제 입장에서 보기에, 코드 자체의 가독성도 준수하게 느껴집니다.)

이 코드를 사용하려면, Api키를 먼저 발급받고 Api키가 들어갈 위치에 Api키를 입력해야 합니다.


```
// onEdit 트리거: A열(1번 열)에 질문이 입력되면 같은 행의 B열(2번 열)에 정적인 GPT 응답을 기록합니다.
function onEdit(e) {
  var sheet = e.range.getSheet();
  var editedRange = e.range;
  
  // 편집된 셀이 A열에 있을 때만 실행
  if (editedRange.getColumn() !== 1) return;
  
  var numRows = editedRange.getNumRows();
  var startRow = editedRange.getRow();
  var values = editedRange.getValues();
  
  for (var i = 0; i < numRows; i++) {
    var prompt = values[i][0];
    // 같은 행의 B열에 이미 값이 있다면 건너뜁니다.
    var answerCell = sheet.getRange(startRow + i, 2);
    if (prompt && answerCell.getValue() === "") {
      var answer = GPT(prompt);
      answerCell.setValue(answer);
      Utilities.sleep(1000); // API 호출 속도 제한 고려
    }
  }
}

// GPT 함수: 주어진 질문(prompt)에 대해 GPT-4o-mini 모델을 호출하여 응답을 반환합니다.
function GPT(prompt) {
  var apiKey = "여기에 Api키 입력";  // 각자의 API 키로 교체하세요.
  if (!prompt || prompt.toString().trim() === "") return "";
  
  var url = "https://api.openai.com/v1/chat/completions";
  var payload = {
    "model": "gpt-4o-mini",
    "messages": [
      { "role": "system", "content": "You are a bot that extracts main word keywords for product reviews." },
      { "role": "user", "content": prompt }
    ],
    "max_tokens": 200
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
    var code = response.getResponseCode();
    if (code !== 200) {
      return "Error: " + code + " " + response.getContentText();
    }
    var json = JSON.parse(response.getContentText());
    if (!json.choices || json.choices.length === 0) {
      return "Error: No response from GPT.";
    }
    return json.choices[0].message.content.trim();
  } catch (e) {
    return "Error: " + e.toString();
  }
}
