package com.irene.app.sharepoint.service;

import java.io.*;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.text.DecimalFormat;
import java.util.Arrays;

import org.apache.http.HttpHost;
import org.apache.http.HttpResponse;
import org.apache.http.client.*;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.*;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.util.EntityUtils;
import org.json.*;
import org.slf4j.*;
import org.springframework.beans.factory.annotation.*;
import org.springframework.http.*;
import org.springframework.lang.Nullable;
import org.springframework.stereotype.Service;
import org.springframework.util.*;
import org.springframework.web.client.RestTemplate;

import com.fasterxml.jackson.databind.JsonNode;

/**
 * @author Irene Hermosilla
 * https://www.anexinet.com/blog/getting-an-access-token-for-sharepoint-online/
 */
@Service
public class SharepointService {
	
	private final Logger log = LoggerFactory.getLogger(this.getClass());
	private static final String APPLICATION_JSON = "application/json;odata=verbose";
	private static final int CHUNK_SIZE = 1024 * 1024 * 200; //200MB
	
	@Value("${app.sharepoint.tenant-name}")
	private String tenantName;
	@Value("${app.sharepoint.tenant-id}")
	private String tenantId;
	@Value("${app.sharepoint.client-id}")
	private String clientId;
	@Value("${app.sharepoint.client-secret}")
	private String clientSecret;
	@Value("${server.proxy.url}")
	private String proxyUrl;
	@Value("${server.proxy.port}")
	private int proxyPort;
	@Value("${server.proxy.default-scheme-name}")
	private String proxyScheme;
	
	@Autowired
	private RestTemplate restTemplate;
	
	public String getToken() {

		String accessToken = "";
		String url = "https://accounts.accesscontrol.windows.net/" + tenantId + "/tokens/OAuth/2";
			
		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(MediaType.APPLICATION_FORM_URLENCODED);
			
		StringBuilder resource = new StringBuilder("00000003-0000-0ff1-ce00-000000000000")
			.append("/")
			.append(tenantName)
			.append(".sharepoint.com")
			.append("@")
			.append(tenantId);
			
		MultiValueMap<String, String> body = new LinkedMultiValueMap<>();
		body.add("grant_type", "client_credentials");
		body.add("client_id", clientId+"@"+tenantId);
		body.add("client_secret", clientSecret);
		body.add("resource", resource.toString());
		
		HttpEntity<MultiValueMap<String, String>> request = new HttpEntity<>(body, headers);
		ResponseEntity<JsonNode> response = restTemplate.postForEntity(url, request, JsonNode.class);
			
		if (response != null && response.getStatusCodeValue() == 200) {
			accessToken = response.getBody().findValue("access_token").textValue();
		} else {
			log.info("Error getting the Sharepoint token: '{}'", response.getStatusCode().toString());
		} 	
		return accessToken;
	}
	
	public void createFolder(String serverRelativeUrl, String token) {
		String url = String.format("https://s%.sharepoint.com/_api/web/folders", tenantName);
		//Maybe could fail - /teams/group/_api
		try {
			String serverUrlEncoded = URLEncoder.encode(serverRelativeUrl, StandardCharsets.UTF_8.toString()).replace("+", "%20");
		
			JSONObject body = new JSONObject();
			body.put("__metadata", new JSONObject().put("type", "SP.Folder"));
			body.put("ServerRelativeUrl", serverUrlEncoded); //Maybe could fail - /teams/group/
		
			HttpHeaders headers = new HttpHeaders();
			headers.add("Authorization", "Bearer "+token);
			headers.add("Accept", APPLICATION_JSON);
			headers.add("Content-Type", APPLICATION_JSON);
		
			HttpEntity<String> request = new HttpEntity<>(body.toString(), headers);
			ResponseEntity<String> response = restTemplate.postForEntity(url, request, String.class);
		
			if (response != null && response.getStatusCodeValue() != 201) {
				log.warn("Error creating folder: '{}'", response.getStatusCode().toString());
			}
		} catch (Exception e) {
			log.warn("Error creating folder: '{}'", e.getMessage());
		}
		
	}
	
	public void uploadFile(String serverRelativeUrl, File file, String uploadFailedPath) {
		String url = "https://s%.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('%s')" 
				+ "/Files/add(url='%s',overwrite=true)"; //Maybe could fail - /teams/group/_api
		try {
			String serverUrlEncoded = URLEncoder.encode(serverRelativeUrl, StandardCharsets.UTF_8.toString()).replace("+", "%20");
			String fileNameEncoded = URLEncoder.encode(file.getName(), StandardCharsets.UTF_8.toString()).replace("+", "%20");
			
			String fileCollectionEndPoint = String.format(url, tenantName, serverUrlEncoded, fileNameEncoded);
			FileEntity entity = new FileEntity(file, ContentType.APPLICATION_OCTET_STREAM);
			HttpResponse response =  executePost(fileCollectionEndPoint, entity);
			
	        if(response != null && response.getStatusLine().getStatusCode() != 200){
				log.warn("Error uploading file: '{}'", response);
				Files.move(file.toPath(), Paths.get(uploadFailedPath+file.getName()), StandardCopyOption.REPLACE_EXISTING);
			}
	        
		} catch (IOException e) {
			log.error("Error sending upload request: '{}'", e.getMessage());
			
			try {
				Files.move(file.toPath(), Paths.get(uploadFailedPath+file.getName()), StandardCopyOption.REPLACE_EXISTING);
			} catch (IOException e1) {
				log.error("Error moving the file to the failed folder: '{}'", e1.getMessage());
			}
		}
		
	}
	
	public void uploadChunkedFile(String serverRelativeUrl, File file, String uploadFailedPath) throws JSONException {
		String baseUrl = "https://%s.sharepoint.com/_api/web/GetFileByServerRelativeUrl('/%s/%s')/%s";
		//Maybe could fail - /teams/group/_api || /teams/group/%s
		
		try {
			String serverUrlEncoded = URLEncoder.encode(serverRelativeUrl, StandardCharsets.UTF_8.toString()).replace("+", "%20").replace("%2F", "/");
			String fileNameEncoded = URLEncoder.encode(file.getName(), StandardCharsets.UTF_8.toString()).replace("+", "%20").replace("%2F", "/");
			
			long fileSize = file.length();
	        if (fileSize <= CHUNK_SIZE) {
	        	uploadFile(serverRelativeUrl, file, uploadFailedPath);
	        	return;
	        } 
	        
			String guid = createDummyFile(serverUrlEncoded, fileNameEncoded);
			String url = "";
			boolean firstChunk = true;
			FileInputStream inputStream = new FileInputStream(file);

			byte[] buffer = new byte[CHUNK_SIZE];
			int read = 0;
			long offset = 0L;
			while ((read = inputStream.read(buffer, 0, buffer.length)) > 0) {

				if (firstChunk) {
					url = String.format(baseUrl, serverUrlEncoded, fileNameEncoded, "StartUpload(uploadId=guid'"+guid+"')"); 
					executeMultiPartRequest(url, buffer, file, uploadFailedPath);
					firstChunk = false;
				} else if (inputStream.available() == 0) {
					url = String.format(baseUrl, serverUrlEncoded, fileNameEncoded, "FinishUpload(uploadId=guid'"+guid+"',fileOffset="+offset+")");
					byte[] finalBuffer = Arrays.copyOf(buffer, read);
					executeMultiPartRequest(url, finalBuffer, file, uploadFailedPath);
				} else {
					url = String.format(baseUrl, serverUrlEncoded, fileNameEncoded, "ContinueUpload(uploadId=guid'"+guid+"',fileOffset="+offset+")");
					executeMultiPartRequest(url, buffer, file, uploadFailedPath);
				}
					
				offset += read;
				float temp = ((float)offset/(float)fileSize) * 100F;
            	log.info("{}% completed", new DecimalFormat("#.##").format(temp));
			}
			inputStream.close();
		} catch (IOException e) {
			log.error("Error sending upload request: '{}'", e.getMessage());
			
			try {
				Files.move(file.toPath(), Paths.get(uploadFailedPath+file.getName()), StandardCopyOption.REPLACE_EXISTING);
			} catch (IOException e1) {
				log.error("Error moving the file to the failed folder: '{}'", e1.getMessage());
			}
		}
		
	}
	
	private String createDummyFile(String serverRelativeUrl, String fileName) throws IOException, JSONException {
        String url = "https://s%.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('%s')"
        		+ "/Files/add(url='%s',overwrite=true)"; //Maybe could fail - /teams/group
        String uniqueId = "";
        
		String fileCollectionEndPoint = String.format(url, serverRelativeUrl, fileName);
		HttpResponse response =  executePost(fileCollectionEndPoint, null);
		
		if (response.getStatusLine().getStatusCode() == 200 || response.getStatusLine().getStatusCode() == 201) {
            String responseString = EntityUtils.toString(response.getEntity(), "UTF-8");
            JSONObject json = new JSONObject(responseString);
            uniqueId = json.getJSONObject("d").getString("UniqueId");
        }
        return uniqueId;
    }
	
	private void executeMultiPartRequest(String url, byte[] fileByteArray, File file, String uploadFailedPath) throws IOException {
		HttpResponse response =  executePost(url, new ByteArrayEntity(fileByteArray));
        
		log.info("Response code: {}. Response: {}", response.getStatusLine().getStatusCode(), response.getStatusLine().getReasonPhrase());
		
		if(response.getStatusLine().getStatusCode() != 200 && Paths.get(uploadFailedPath+file.getName()).toFile().exists()){
			log.warn("Error uploading file: '{}'", response);
			Files.move(file.toPath(), Paths.get(uploadFailedPath+file.getName()), StandardCopyOption.REPLACE_EXISTING);
		}
    }
	
	private HttpResponse executePost(String url, @Nullable AbstractHttpEntity entity) throws IOException {
        HttpPost post = new HttpPost(url);
        post.addHeader("Accept", APPLICATION_JSON);
        post.setHeader("Content-Type", APPLICATION_JSON);
        post.setHeader("Authorization", "Bearer " + getToken());
        if(entity != null){  post.setEntity(entity);  }
       
        HttpHost proxy = new HttpHost(proxyUrl, proxyPort, proxyScheme);
        HttpClient client = HttpClientBuilder.create().setProxy(proxy).build();
        return client.execute(post);
    }
	
}
