package pe.avancecampo.app;

import android.app.Activity;
import android.os.Bundle;
import android.provider.Settings;
import android.webkit.JavascriptInterface;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;

import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;

public class MainActivity extends Activity {
    private WebView web;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        web = new WebView(this);
        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setDatabaseEnabled(true);
        s.setAllowFileAccess(true);
        web.setWebViewClient(new WebViewClient());
        web.addJavascriptInterface(new Bridge(), "AndroidApp");
        setContentView(web);
        if (savedInstanceState != null) {
            web.restoreState(savedInstanceState);
        } else {
            web.loadUrl("file:///android_asset/index.html");
        }
    }

    @Override
    protected void onResume() {
        super.onResume();
        if (web != null) web.evaluateJavascript("window.onAppResume && window.onAppResume()", null);
    }

    @Override
    protected void onSaveInstanceState(Bundle outState) {
        super.onSaveInstanceState(outState);
        web.saveState(outState);
    }

    @Override
    public void onBackPressed() {
        web.evaluateJavascript("window.onBack ? window.onBack() : false", value -> {
            if (!"true".equals(value)) {
                if (web.canGoBack()) web.goBack();
                else MainActivity.super.onBackPressed();
            }
        });
    }

    /** GET con redirecciones manuales (Google Apps Script responde con 302). */
    private static String httpGet(String address) throws Exception {
        String current = address;
        for (int i = 0; i < 6; i++) {
            HttpURLConnection c = (HttpURLConnection) new URL(current).openConnection();
            c.setInstanceFollowRedirects(false);
            c.setConnectTimeout(15000);
            c.setReadTimeout(25000);
            c.setRequestMethod("GET");
            int code = c.getResponseCode();
            if (code >= 300 && code < 400) {
                String loc = c.getHeaderField("Location");
                c.disconnect();
                if (loc == null) throw new Exception("redireccion");
                current = new URL(new URL(current), loc).toString();
                continue;
            }
            InputStream in = code >= 400 ? c.getErrorStream() : c.getInputStream();
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            if (in != null) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) > 0) out.write(buf, 0, n);
                in.close();
            }
            c.disconnect();
            if (code >= 400) throw new Exception("HTTP " + code);
            return out.toString("UTF-8");
        }
        throw new Exception("demasiadas redirecciones");
    }

    public class Bridge {
        @JavascriptInterface
        public String deviceId() {
            String id = Settings.Secure.getString(getContentResolver(), Settings.Secure.ANDROID_ID);
            return id == null ? "" : id;
        }

        @JavascriptInterface
        public void http(final String callId, final String url) {
            new Thread(() -> {
                boolean ok;
                String text;
                try {
                    text = httpGet(url);
                    ok = true;
                } catch (Exception e) {
                    text = String.valueOf(e.getMessage());
                    ok = false;
                }
                final String js = "window.__http(" + JSONObject.quote(callId) + "," + ok + "," + JSONObject.quote(text) + ")";
                runOnUiThread(() -> web.evaluateJavascript(js, null));
            }).start();
        }
    }
}
