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
			mailSender.setHost("mail.id4-realms.com");
			mailSender.setPort(587);
			mailSender.setUsername("vishal.jaiswal@id4-realms.com");
			mailSender.setPassword("id4@123A");

			File file = new File("D://AWS//production-report.xlsx");

			InputStream targetStream = null;
			targetStream = new FileInputStream(file);

			String html = new ExcelToHtml(targetStream).getHTML();

			String path = "D:/AWS/text.txt";
			Files.write(Paths.get(path), html.getBytes());

			MimeMessage message = mailSender.createMimeMessage();
			MimeMessageHelper helper = new MimeMessageHelper(message, true);
			helper.setTo("svgurcharan@gmail.com");
			helper.setFrom("vishal.jaiswal@id4-realms.com");
			helper.setText(html, true);

			/*
			 * helper.addAttachment(attachFile.getOriginalFilename(), new
			 * InputStreamSource() {
			 * 
			 * @Override public InputStream getInputStream() throws IOException { return
			 * attachFile.getInputStream(); } });
			 */
			helper.setSubject("Report with attachment");

			mailSender.send(message);
			
			return "success";
		} catch (Exception e) {
			e.printStackTrace();
		}

		return "failed";
	}

}
