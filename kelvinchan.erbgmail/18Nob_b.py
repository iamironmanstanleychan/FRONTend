<!DOCTYPE html>
<html lang="zh-HK">

<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>留言板：新增、插入、替換與刪除</title>
    <style>
        * {
            box-sizing: border-box;
        }

        body {
            margin: 0;
            padding: 24px;
            background: #f4f6fa;
            color: #1f2937;
            font-family: system-ui, sans-serif;
            line-height: 1.6;
        }

        main {
            max-width: 720px;
            margin: 0 auto;
            padding: 24px;
            border-radius: 12px;
            background: white;
        }

        h1 {
            margin-top: 0;
            font-size: 1.6rem;
        }

        label {
            display: block;
            font-weight: bold;
        }

        textarea {
            width: 100%;
            padding: 10px;
            border: 1px solid #64748b;
            border-radius: 6px;
            font: inherit;
            resize: vertical;
        }

        .buttons {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-top: 12px;
        }

        button {
            padding: 10px 14px;
            border: 0;
            border-radius: 6px;
            background: #1d4ed8;
            color: white;
            font: inherit;
            cursor: pointer;
        }

        #noticeBtn {
            background: #854d0e;
        }

        button:focus-visible,
        textarea:focus-visible {
            outline: 3px solid #0369a1;
            outline-offset: 3px;
        }

        #messageList {
            padding: 0;
            list-style: none;
        }

        #messageList li {
            margin: 10px 0;
            padding: 12px;
            border-left: 4px solid #1d4ed8;
            background: #eff6ff;
            overflow-wrap: anywhere;
        }

        .message-text {
            display: block;
            white-space: pre-wrap;
        }

        .message-actions {
            display: flex;
            gap: 8px;
            margin-top: 10px;
        }

        .message-actions button {
            padding: 6px 12px;
        }

        .edit-btn {
            background: #047857;
        }

        .delete-btn {
            background: #b91c1c;
        }

        #messageList .notice {
            border-color: #a16207;
            background: #fef3c7;
        }

        pre {
            padding: 12px;
            background: #f1f5f9;
            white-space: pre-wrap;
            overflow-wrap: anywhere;
        }
    </style>
</head>

<body>
    <main>
        <h1>班級留言板</h1>
        <p>新留言加到最尾；新公告插到列表最前面。</p>
        <p>每則留言亦可以編輯或刪除，示範替換及移除子節點。</p>

        <label for="messageInput">留言／公告內容</label>
        <textarea id="messageInput" rows="2" placeholder="例如：聽日記得帶手提電腦。"></textarea>

        <div class="buttons">
            <button type="button" id="sendBtn">發佈留言（加到最尾）</button>
            <button type="button" id="noticeBtn">發佈公告（插到最前）</button>
        </div>

        <p id="feedback" role="status">輸入內容，再選擇其中一個按鈕。</p>

        <h2>留言列表</h2>
        <ul id="messageList"></ul>

        <h2>對應程式碼</h2>

        <p><strong>appendChild()：</strong>加入父節點嘅子節點列表末尾。</p>
        <pre><code>messageList.appendChild(newMessage);</code></pre>

        <p><strong>insertBefore()：</strong>插入指定子節點之前，呢度選擇第一個元素。</p>
        <pre><code>messageList.insertBefore(
  newMessage,
  messageList.firstElementChild
);</code></pre>

        <p><strong>replaceChild()：</strong>用新留言元素替換舊留言元素。</p>
        <pre><code>messageList.replaceChild(newMessage, message);
// 第一個參數：新節點；第二個參數：舊節點。</code></pre>

        <p><strong>removeChild()：</strong>移除指定留言元素。</p>
        <pre><code>messageList.removeChild(message);</code></pre>

        <p><small>呢個示範只改動目前頁面，重新整理後會還原。</small></p>
    </main>

    <script>
        const messageList = document.querySelector("#messageList");
        const messageInput = document.querySelector("#messageInput");
        const feedback = document.querySelector("#feedback");

        // 建立完整留言：文字 + 編輯按鈕 + 刪除按鈕。
        function createMessage(text, prefix = "我：", isNotice = false) {
            const message = document.createElement("li");
            if (isNotice) message.classList.add("notice");

            const content = document.createElement("span");
            content.className = "message-text";
            content.textContent = prefix + text;

            const actions = document.createElement("div");
            actions.className = "message-actions";

            const editButton = document.createElement("button");
            editButton.type = "button";
            editButton.className = "edit-btn";
            editButton.textContent = "編輯";

            const deleteButton = document.createElement("button");
            deleteButton.type = "button";
            deleteButton.className = "delete-btn";
            deleteButton.textContent = "刪除";

            // 3. 編輯：先建立新 li，再替換原本嘅 li。
            editButton.addEventListener("click", () => {
                const updatedText = prompt("請輸入修改後嘅內容：", text);

                // 按「取消」時，保留原本留言。
                if (updatedText === null) return;

                if (updatedText.trim() === "") {
                    feedback.textContent = "修改內容唔可以留空。";
                    editButton.focus();
                    return;
                }

                const newMessage = createMessage(updatedText.trim(), prefix, isNotice);

                // 次序：新節點先，舊節點後。
                messageList.replaceChild(newMessage, message);

                feedback.textContent = "已用 replaceChild() 替換整個留言元素。";
                newMessage.querySelector(".edit-btn").focus();
            });

            // 4. 刪除：由父節點移除指定嘅 li。
            deleteButton.addEventListener("click", () => {
                messageList.removeChild(message);

                feedback.textContent = "已用 removeChild() 刪除呢則留言。";
                messageInput.focus();
            });

            actions.appendChild(editButton);
            actions.appendChild(deleteButton);
            message.appendChild(content);
            message.appendChild(actions);

            return message;
        }

        // 共用函式：取得輸入內容，並阻止空白留言。
        function getInputText() {
            const text = messageInput.value.trim();

            if (text === "") {
                feedback.textContent = "請先輸入內容。";
                messageInput.focus();
                return null;
            }

            messageInput.value = "";
            messageInput.focus();
            return text;
        }

        // 1. 發佈普通留言：加到列表最尾。
        document.querySelector("#sendBtn").addEventListener("click", () => {
            const text = getInputText();
            if (text === null) return;

            const newMessage = createMessage(text);

            messageList.appendChild(newMessage);

            feedback.textContent = "已用 appendChild() 將留言加到最尾。";
        });

        // 2. 發佈公告：插到列表第一個元素之前。
        document.querySelector("#noticeBtn").addEventListener("click", () => {
            const text = getInputText();
            if (text === null) return;

            const newMessage = createMessage(text, "【公告】", true);

            messageList.insertBefore(
                newMessage,
                messageList.firstElementChild
            );

            feedback.textContent = "已用 insertBefore() 將公告插到最前。";
        });

        // 原有留言亦用同一個函式建立，所以都有編輯及刪除按鈕。
        messageList.appendChild(createMessage("今日開始學習 DOM！", "阿明："));
        messageList.appendChild(createMessage("試試用 JavaScript 新增留言。", "阿欣："));
    </script>
</body>

</html>