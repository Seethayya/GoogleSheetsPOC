package com.googlesheets.poc.controller;

import java.util.Calendar;
import java.util.List;

import javax.servlet.http.HttpSession;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import com.googlesheets.poc.model.GAuthority;
import com.googlesheets.poc.model.GSheet;
import com.googlesheets.poc.model.GToken;
import com.googlesheets.poc.model.Response;
import com.googlesheets.poc.service.GoogleApiService;
import com.googlesheets.poc.service.GoogleApplicationService;

@RestController
@RequestMapping("/api")
public class GoogleSheetsController {

	private static final Logger LOGGER = Logger.getLogger(GoogleSheetsController.class);
	@Autowired
	private GoogleApiService googleApiService;
	@Autowired
	private GoogleApplicationService applicationService;
	
	@RequestMapping(value = "/findCustomer/{email}", method = RequestMethod.GET)
	@ResponseBody
    public Response findCustomer(@PathVariable String email, HttpSession session) {
		LOGGER.debug("findCustomer()"+email);
		Response response = new Response();
		try {
			session.setAttribute("Email_Id", email);
			GAuthority token = (GAuthority) session.getAttribute(email);
			if (token == null) {
				response.setRedirectUril(applicationService.redirectUril());
				response.setStatus("NEED_AUTHENTICATION");
			} else {
				response.setStatus("Already authenticated");
			}
			//String token = applicationService.getTokenByRefresh("1/aL1IOJ4WV7gbFuinIpHvmUGEirFKukeAA1JT04M-YAFQmxy4i_k8YXHIWO32PyLm");
			
			//sheets = googleApiService.findSheet("1MqAFHkTtzcH9h8Rup5lCDVKAQBmLvbC45-udQlBHbgM", token);;
		} catch (Exception e) {
			LOGGER.error(e);
			response.setStatus("Failed "+e.getMessage());
		}
        return response;
    }
	
	@RequestMapping(value = "/generateToken", method = RequestMethod.POST)
	@ResponseBody
    public Response generateToken(@RequestBody String code, HttpSession session) {
		LOGGER.debug("generateToken()"+code);
		Response response = new Response();
		List<GSheet> sheets = null;
		String email =	(String) session.getAttribute("Email_Id");
		response.setCustomerName(email);
		try {
			GAuthority gAuthority = new GAuthority();	
			GToken token = applicationService.getTokenByCode(code);
			System.out.println(token);
			if (token != null) {
				gAuthority.setAccessToken(token.getAccess_token());
				if (StringUtils.isNotBlank(token.getExpires_in()))
					gAuthority.setExpires(Integer.valueOf(token.getExpires_in()));
				gAuthority.setTokenType(token.getToken_type());
				gAuthority.setCreatedTime(Calendar.getInstance().getTime());
				gAuthority.setRefreshToken(token.getRefresh_token());
				session.setAttribute(email, gAuthority);
				response.setStatus("Success");
			} else {
				response.setStatus("Failed");
			}
		} catch (Exception e) {
			LOGGER.error(e);
		}
		response.setSheets(sheets);
		response.setStatus("Success");
        return response;
    }
	
	@RequestMapping(value = "/getSpreadSheets/{email}/{spreadSheetId}", method = RequestMethod.GET)
	@ResponseBody
    public Response getSpreadSheets(@PathVariable String email, @PathVariable String spreadSheetId, HttpSession session) {
		LOGGER.debug("getSpreadSheets email"+email+" spreadSheetId:"+spreadSheetId);
		Response response = new Response();
		try {
			GAuthority gAuthority = (GAuthority) session.getAttribute(email);
			if (gAuthority == null) {
				response.setStatus("Session Expired");
				return response;
			}
			if (StringUtils.isBlank(gAuthority.getRefreshToken())) {
				long currentTime = Calendar.getInstance().getTime().getTime()/(1000 * 60);
				long createdTime = gAuthority.getCreatedTime().getTime()/(1000 * 60);
				//Added buffer 10min
				if (currentTime - createdTime > gAuthority.getExpires()-10) {
					LOGGER.debug("Token expired");
					GToken token = applicationService.getTokenByRefresh(gAuthority.getRefreshToken());
					if (token != null) {
						gAuthority.setAccessToken(token.getAccess_token());
						if (StringUtils.isNotBlank(token.getExpires_in()))
							gAuthority.setExpires(Integer.valueOf(token.getExpires_in()));
						gAuthority.setTokenType(token.getToken_type());
						gAuthority.setCreatedTime(Calendar.getInstance().getTime());
						session.setAttribute(email, gAuthority);
					} else {
						response.setStatus("Autherization Problem");
						return response;
					}
				}
			}
			if (StringUtils.isBlank(gAuthority.getAccessToken())) {
				response.setStatus("Autherization Problem");
				return response;
			}
			List<GSheet> sheets = googleApiService.findSheet(spreadSheetId, gAuthority.getAccessToken());
			response.setSheets(sheets);
			response.setStatus("Success");
		} catch (Exception e) {
			LOGGER.error(e);
			response.setStatus("No Sheets available");
		}
        return response;
    }
}
