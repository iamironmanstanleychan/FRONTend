<!DOCTYPE html>
<html lang="zh-HK">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>DOM 遍歷：社交平台動態牆</title>

    <style>
        * {
            box-sizing: border-box;
        }

        body {
            margin: 0;
            padding: 20px;
            background: #f0f2f5;
            font-family: Arial, sans-serif;
            color: #222;
        }

        .container {
            max-width: 650px;
            margin: 0 auto;
        }

        h1 {
            text-align: center;
            color: #1877f2;
        }

        .toolbar {
            background: white;
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 20px;
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }

        button {
            padding: 9px 12px;
            border: 0;
            border-radius: 6px;
            background: #e4e6eb;
            cursor: pointer;
            font-size: 14px;
        }

        button:hover {
            background: #d8dadf;
        }

        .toolbar button {
            background: #1877f2;
            color: white;
        }

        .post {
            background: white;
            padding: 20px;
            margin-bottom: 15px;
            border: 3px solid transparent;
            border-radius: 10px;
            transition: border-color 0.2s;
        }

        .post.selected {
            border-color: #1877f2;
        }

        .post.liked {
            background: #eef5ff;
        }

        .author {
            font-weight: bold;
            margin-bottom: 12px;
        }

        .content {
            line-height: 1.6;
        }

        .actions {
            margin-top: 15px;
            padding-top: 12px;
            border-top: 1px solid #ddd;
            display: flex;
            gap: 8px;
        }

        .liked .like-btn {
            background: #1877f2;
            color: white;
        }

        #status {
            background: #fff8db;
            padding: 15px;
            margin-top: 20px;
            border-radius: 8px;
            line-height: 1.6;
        }
    </style>
</head>

<body>
    <main class="container">
        <h1>社交平台動態牆</h1>

        <div class="toolbar">
            <button id="firstBtn">第一篇</button>
            <button id="lastBtn">最後一篇</button>
            <button id="prevBtn">上一篇</button>
            <button id="nextBtn">下一篇</button>
            <button id="allBtn">顯示所有貼文</button>
        </div>

        <!-- 所有貼文都放在同一個父元素 -->
        <section id="feed" aria-label="貼文動態牆">

            <article class="post">
                <div class="author">陳大文</div>
                <p class="content">今日去咗行山，天氣真係好好！☀️</p>
                <div class="actions">
                    <button class="like-btn">👍 讚好</button>
                </div>
            </article>

            <article class="post">
                <div class="author">李小明</div>
                <p class="content">最近開始學 JavaScript，DOM 好有趣！💻</p>
                <div class="actions">
                    <button class="like-btn">👍 讚好</button>
                </div>
            </article>

            <article class="post">
                <div class="author">黃美玲</div>
                <p class="content">今晚食咗一餐好豐富嘅晚飯！🍜</p>
                <div class="actions">
                    <button class="like-btn">👍 讚好</button>
                </div>
            </article>

            <article class="post">
                <div class="author">張志強</div>
                <p class="content">星期六有冇人一齊打羽毛球？🏸</p>
                <div class="actions">
                    <button class="like-btn">👍 讚好</button>
                </div>
            </article>

            <article class="post">
                <div class="author">何嘉欣</div>
                <p class="content">終於完成咗自己第一個網站！🎉</p>
                <div class="actions">
                    <button class="like-btn">👍 讚好</button>
                </div>
            </article>

        </section>

        <div id="status" role="status" aria-live="polite">
            請按上方按鈕操作貼文。
        </div>
    </main>

    <script>
        const feed = document.querySelector("#feed");
        const status = document.querySelector("#status");

        // 記錄目前選取的貼文
        let currentPost = null;

        // 共用函式：選取並突出顯示貼文
        function selectPost(post) {
            if (!post) {
                status.textContent = "找不到貼文！";
                return;
            }

            // 移除之前選取的效果
            if (currentPost) {
                currentPost.classList.remove("selected");
            }

            // 選取新貼文
            currentPost = post;
            currentPost.classList.add("selected");

            const author = currentPost.querySelector(".author");
            status.textContent = "目前選取：" + author.textContent;

            currentPost.scrollIntoView({
                behavior: "smooth",
                block: "center"
            });
        }

        // ====================================
        // 1. firstElementChild
        // 選取第一篇貼文
        // ====================================
        document.querySelector("#firstBtn")
            .addEventListener("click", function () {

                const firstPost = feed.firstElementChild;

                selectPost(firstPost);
            });

        // ====================================
        // 2. lastElementChild
        // 選取最後一篇貼文
        // ====================================
        document.querySelector("#lastBtn")
            .addEventListener("click", function () {

                const lastPost = feed.lastElementChild;

                selectPost(lastPost);
            });

        // ====================================
        // 3. previousElementSibling
        // 選取上一篇貼文
        // ====================================
        document.querySelector("#prevBtn")
            .addEventListener("click", function () {

                if (!currentPost) {
                    selectPost(feed.firstElementChild);
                    return;
                }

                const previousPost =
                    currentPost.previousElementSibling;

                if (previousPost) {
                    selectPost(previousPost);
                } else {
                    status.textContent = "已經係第一篇貼文！";
                }
            });

        // ====================================
        // 4. nextElementSibling
        // 選取下一篇貼文
        // ====================================
        document.querySelector("#nextBtn")
            .addEventListener("click", function () {

                if (!currentPost) {
                    selectPost(feed.firstElementChild);
                    return;
                }

                const nextPost =
                    currentPost.nextElementSibling;

                if (nextPost) {
                    selectPost(nextPost);
                } else {
                    status.textContent = "已經係最後一篇貼文！";
                }
            });

        // ====================================
        // 5. children
        // 遍歷所有貼文
        // ====================================
        document.querySelector("#allBtn")
            .addEventListener("click", function () {

                const posts = feed.children;

                let result = "共有 " + posts.length + " 篇貼文：\n";

                for (const post of posts) {
                    const author = post.querySelector(".author");
                    result += author.textContent + "、";
                }

                status.textContent = result;
            });

        // ====================================
        // 6. parentNode
        // 由讚好按鈕向上尋找所屬貼文
        // ====================================
        const likeButtons = document.querySelectorAll(".like-btn");

        for (const button of likeButtons) {
            button.addEventListener("click", function () {

                // button → div.actions
                const actions = this.parentNode;

                // div.actions → article.post
                const post = actions.parentNode;

                // 切換讚好狀態
                post.classList.toggle("liked");

                // 更新按鈕文字
                if (post.classList.contains("liked")) {
                    this.textContent = "💙 已讚好";
                } else {
                    this.textContent = "👍 讚好";
                }

                const author = post.querySelector(".author");

                status.textContent =
                    "你操作咗 " + author.textContent + " 嘅貼文";
            });
        }
    </script>
</body>

</html>