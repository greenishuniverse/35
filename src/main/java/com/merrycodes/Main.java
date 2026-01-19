package com.merrycodes;

import com.microsoft.graph.auth.confidentialClient.ClientCredentialProvider;
import com.microsoft.graph.auth.enums.NationalCloud;
import com.microsoft.graph.models.extensions.IGraphServiceClient;
import com.microsoft.graph.models.extensions.Message;
import com.microsoft.graph.models.extensions.User;
import com.microsoft.graph.requests.extensions.GraphServiceClient;
import com.microsoft.graph.requests.extensions.IMessageCollectionPage;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Collections;
import java.util.List;

/**
 * Github Action 版本 <br>
 * Office E5 调用 graph API <br>
 */
public class Main {

    public static void main(String[] args) {
        // args 0 = CLIENT_ID
        // args 1 = TENANT_GUID
        // args 2 = CLIENT_SECRET
        // 旧的 USERNAME/PASSWORD 坐标在 args[3], args[4] 但不再使用

        ClientCredentialProvider authProvider = new ClientCredentialProvider(
                args[0],
                Collections.singletonList("https://graph.microsoft.com/.default"),
                args[2],
                args[1]
        );

        IGraphServiceClient graphClient = GraphServiceClient.builder()
                .authenticationProvider(authProvider)
                .buildClient();

        // App-only 模式下没有 /me，所以我们换一个 API 调用
        // 比如列出用户的邮件（等价于原来你用 /me 列邮件）
        List<User> users = graphClient.users().buildRequest().top(5).get().getCurrentPage();

        System.out.println(String.format("运行时间：%s —— 共有%d 名用户",
                LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")),
                users.size()));

        // 示例：遍历几个用户并尝试列邮件
        for (User u : users) {
            System.out.println("User: " + u.userPrincipalName);

            IMessageCollectionPage mailPage = graphClient.users(u.id)
                    .messages()
                    .buildRequest()
                    .select("sender,subject")
                    .top(5)
                    .get();

            List<Message> messageList = mailPage.getCurrentPage();
            System.out.println("  " + u.userPrincipalName + " 有 " + messageList.size() + " 封邮件");
        }
    }
}
