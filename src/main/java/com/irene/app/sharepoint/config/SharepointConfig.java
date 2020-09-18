package com.irene.app.sharepoint.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.client.RestTemplate;

/**
 * @author Irene Hermosilla
 */
@Configuration
public class SharepointConfig {
	
	@Bean
	public RestTemplate getRestTemplate() {
	     return new RestTemplate();
	}

}
