package com.googlesheets.poc.service;

import javax.annotation.PostConstruct;

import org.apache.log4j.Logger;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.ResponseEntity;
import org.springframework.http.client.HttpComponentsClientHttpRequestFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.util.UriComponentsBuilder;

import com.google.api.services.sheets.v4.SheetsScopes;
import com.googlesheets.poc.model.GApplicationCredential;
import com.googlesheets.poc.model.GToken;

@Service
public class GoogleApplicationService {

    private static final Logger LOGGER = Logger.getLogger(GoogleApplicationService.class);

    
	@Value("${client_id}")
	private String clientId;
	@Value("${project_id}")
	private String projctId;
	@Value("${auth_uri}")
	private String authUri;
	@Value("${token_uri}")
	private String toeknUri;
	@Value("${auth_provider_x509_cert_url}")
	private String athProviderUrl;
	@Value("${client_secret}")
	private String secretId;
	@Value("${redirectUri}")
	private String redirectUri;
	private GApplicationCredential applicationCred;
	
	
	@PostConstruct
	public void constructApplicationCred() {
		applicationCred = new GApplicationCredential(clientId, projctId, authUri, toeknUri, athProviderUrl, secretId, redirectUri);
	}
	
	public static void main(String[] args) {
		System.out.println(new GoogleApplicationService().getTokenByCode("4/S49Ibn08_wWdrLsea0fOulBQlG7v7zu6bOxNTBQnPwc"));
	}
	
	public GToken getTokenByRefresh(String refreshToken) {
		HttpComponentsClientHttpRequestFactory requestFactory = new HttpComponentsClientHttpRequestFactory();
		requestFactory.setConnectTimeout(3 * 10000);
		UriComponentsBuilder builder = UriComponentsBuilder.fromHttpUrl(toeknUri);
		builder.queryParam("client_id", applicationCred.getClientId());
		builder.queryParam("client_secret", applicationCred.getSecretId());
		builder.queryParam("refresh_token", refreshToken);
		builder.queryParam("grant_type", "refresh_token");
        System.out.println("---builder:"+builder.toUriString());
		RestTemplate rt = new RestTemplate(requestFactory);
		GToken gtoken = null;
		try {
			gtoken = rt.postForObject(builder.toUriString(), null,GToken.class);
		} catch (Exception e) {
			LOGGER.error(e);
		}
		System.out.println(gtoken);
		return gtoken;
	}
	
	public GToken getTokenByCode(String code) {
		LOGGER.debug("getTokenByCode"+code);
		HttpComponentsClientHttpRequestFactory requestFactory = new HttpComponentsClientHttpRequestFactory();
		requestFactory.setConnectTimeout(3 * 10000);
		//toeknUri = "https://www.googleapis.com/oauth2/v4/token";
		UriComponentsBuilder builder = UriComponentsBuilder.fromHttpUrl(toeknUri);
		builder.queryParam("code", code);
		builder.queryParam("client_id", applicationCred.getClientId());
		builder.queryParam("client_secret", applicationCred.getSecretId());
		builder.queryParam("redirect_uri", redirectUri);
		builder.queryParam("grant_type", "authorization_code");
		RestTemplate rt = new RestTemplate(requestFactory);
		ResponseEntity<GToken> response = null;
		LOGGER.debug("-----URL:"+builder.toUriString());
		try {
			response = rt.postForEntity(builder.toUriString(), null, GToken.class);
		} catch (Exception e) {
			LOGGER.error(e);
		}
		if (response != null)
			return response.getBody();
		return null;
	}
	
	
	public String redirectUril() {
		UriComponentsBuilder builder = UriComponentsBuilder.fromHttpUrl(authUri);
		builder.queryParam("scope",SheetsScopes.SPREADSHEETS_READONLY);
		builder.queryParam("access_type","offline");
		builder.queryParam("include_granted_scopes","true");
		builder.queryParam("state","state_parameter_passthrough_value");
		builder.queryParam("redirect_uri",redirectUri);
		builder.queryParam("response_type","code");
		builder.queryParam("client_id",clientId);
		return builder.toUriString();
	}
}
