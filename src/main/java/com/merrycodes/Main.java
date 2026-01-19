package com.merrycodes;

import com.microsoft.graph.auth.enums.NationalCloud;
import com.microsoft.graph.auth.publicClient.UsernamePasswordProvider;
import com.microsoft.graph.models.extensions.IGraphServiceClient;
import com.microsoft.graph.models.extensions.Message;
import com.microsoft.graph.models.extensions.User;
import com.microsoft.graph.requests.extensions.GraphServiceClient;
import com.microsoft.graph.requests.extensions.IMessageCollectionPage;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Collections;
import java.util.List;

/**
 * Github Action 版本 <br>
 * Office E5 调用 graph API <br>
 * <a href="https://docs.microsoft.com/zh-cn/graph/api/resources/message?view=graph-rest-1.0">关于邮件的操作</a>
 *
 * @author MerryCodes
 * @date 2020/12/19 16:49
 */
public class Main {

    /**
     * 参数一定要正确，否者会调用失败
     *
     * @param args 参数要求查看README.md
     */
    public static void main(String[] args) {
        if (args == null || (args.length != 4 && args.length != 5)) {
            System.out.println("参数错误。支持两种模式：\n" +
                    "1) 旧版(用户名/密码，可能会被MFA/条件访问阻断)：\n" +
                    "   java -jar ... <CLIENT_ID> <USERNAME> <PASSWORD> <TENANT_GUID> <CLIENT_SECRET>\n" +
                    "2) 推荐(应用模式，无需用户名密码，兼容MFA)：\n" +
                    "   java -jar ... <CLIENT_ID> <TENANT_GUID> <CLIENT_SECRET> <USER_UPN>");
            System.exit(2);
            return;
        }

        // 推荐：应用模式（client_credentials），避免密码变更/MFA导致的失败
        if (args.length == 4) {
            String clientId = args[0];
            String tenantGuid = args[1];
            String clientSecret = args[2];
            String userUpn = args[3];

            String token = fetchClientCredentialToken(tenantGuid, clientId, clientSecret);
            // 用应用权限访问指定用户信息（需要为应用授予 User.Read.All 并管理员同意）
            String encodedUpn = urlEncode(userUpn);
            String url = "https://graph.microsoft.com/v1.0/users/" + encodedUpn + "?$select=id,displayName,userPrincipalName";
            int status = callGraphGet(token, url);
            System.out.println(String.format("运行时间：%s —— Graph调用状态码=%d", LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")), status));
            return;
        }

        // 旧版：用户名/密码（ROPC）。若账号启用MFA/安全默认值/条件访问，可能直接失败。
        UsernamePasswordProvider authProvider = new UsernamePasswordProvider(args[0], Collections.singletonList("https://graph.microsoft.com/.default"),
                args[1], args[2], NationalCloud.Global, args[3], args[4]);
        IGraphServiceClient graphClient = GraphServiceClient.builder().authenticationProvider(authProvider).buildClient();
        User user = graphClient.me().buildRequest().get();
        IMessageCollectionPage iMessageCollectionPage = graphClient.users(user.userPrincipalName).messages().buildRequest().select("sender,subject").get();
        List<Message> messageList = iMessageCollectionPage.getCurrentPage();
        System.out.println(String.format("运行时间：%s —— 共有%d封件邮件", LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")), messageList.size()));
    }

    private static String fetchClientCredentialToken(String tenantGuid, String clientId, String clientSecret) {
        try {
            String tokenUrl = "https://login.microsoftonline.com/" + urlEncode(tenantGuid) + "/oauth2/v2.0/token";
            String body = "client_id=" + urlEncode(clientId)
                    + "&client_secret=" + urlEncode(clientSecret)
                    + "&scope=" + urlEncode("https://graph.microsoft.com/.default")
                    + "&grant_type=client_credentials";

            HttpURLConnection conn = (HttpURLConnection) new URL(tokenUrl).openConnection();
            conn.setRequestMethod("POST");
            conn.setDoOutput(true);
            conn.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");

            try (OutputStream os = conn.getOutputStream()) {
                os.write(body.getBytes(StandardCharsets.UTF_8));
            }

            int code = conn.getResponseCode();
            String resp = readAll(code >= 200 && code < 300 ? conn.getInputStream() : conn.getErrorStream());
            if (code < 200 || code >= 300) {
                throw new RuntimeException("获取token失败，HTTP=" + code + "，响应=" + resp);
            }

            String token = extractJsonString(resp, "access_token");
            if (token == null || token.isEmpty()) {
                throw new RuntimeException("获取token失败：响应中未找到access_token，响应=" + resp);
            }
            return token;
        } catch (IOException e) {
            throw new RuntimeException("获取token失败：" + e.getMessage(), e);
        }
    }

    private static int callGraphGet(String bearerToken, String url) {
        try {
            HttpURLConnection conn = (HttpURLConnection) new URL(url).openConnection();
            conn.setRequestMethod("GET");
            conn.setRequestProperty("Authorization", "Bearer " + bearerToken);
            conn.setRequestProperty("Accept", "application/json");

            int code = conn.getResponseCode();
            // 读一下响应，避免某些运行环境下连接未完全释放
            readAll(code >= 200 && code < 300 ? conn.getInputStream() : conn.getErrorStream());
            return code;
        } catch (IOException e) {
            throw new RuntimeException("Graph请求失败：" + e.getMessage(), e);
        }
    }

    private static String readAll(InputStream is) throws IOException {
        if (is == null) return "";
        try (InputStream in = is; ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            byte[] buf = new byte[4096];
            int n;
            while ((n = in.read(buf)) >= 0) {
                baos.write(buf, 0, n);
            }
            return baos.toString(StandardCharsets.UTF_8.name());
        }
    }

    private static String extractJsonString(String json, String key) {
        if (json == null || key == null) return null;
        String pattern = "\"" + key + "\"";
        int keyIdx = json.indexOf(pattern);
        if (keyIdx < 0) return null;
        int colonIdx = json.indexOf(':', keyIdx + pattern.length());
        if (colonIdx < 0) return null;
        int firstQuote = json.indexOf('"', colonIdx + 1);
        if (firstQuote < 0) return null;
        int secondQuote = json.indexOf('"', firstQuote + 1);
        if (secondQuote < 0) return null;
        return json.substring(firstQuote + 1, secondQuote);
    }

    private static String urlEncode(String s) {
        try {
            return URLEncoder.encode(s, StandardCharsets.UTF_8.name());
        } catch (Exception e) {
            // StandardCharsets.UTF_8 is always supported
            return s;
        }
    }

}
