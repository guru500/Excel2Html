package com.excel2html.converter.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

import javax.mail.internet.MimeMessage;

import org.springframework.mail.javamail.JavaMailSenderImpl;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

/**
 * @author Gurcharan
 *
 */
@RestController
@RequestMapping("/mail")
public class IprodMailController {

	@RequestMapping("home")
	public String home() {
		return "Home Mail !";

	}

	@RequestMapping(value = "/sendMail", method = RequestMethod.POST)
	public String sendMail(@RequestParam("file") MultipartFile attachFile) {

		try {

			JavaMailSenderImpl mailSender = new JavaMailSenderImpl();
			mailSender.setHost("****");
			mailSender.setPort(123);
			mailSender.setUsername("****");
			mailSender.setPassword("******");

			File file = new File("***********");

			InputStream targetStream = null;
			targetStream = new FileInputStream(file);

			String html = new ExcelToHtml(targetStream).getHTML();

			String path = "******";
			Files.write(Paths.get(path), html.getBytes());

			MimeMessage message = mailSender.createMimeMessage();
			MimeMessageHelper helper = new MimeMessageHelper(message, true);
			helper.setTo("****************");
			helper.setFrom("********************");
			helper.setText(html, true);

			helper.setSubject("Report with attachment");

			mailSender.send(message);
			
			return "success";
		} catch (Exception e) {
			e.printStackTrace();
		}

		return "failed";
	}

}
