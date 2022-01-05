package com.kafka.demo.controller;

import java.util.Date;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
public class WordController {
	
	
	@GetMapping("/word")
	public String sayHello(Model themodel) {
		themodel.addAttribute("theDate", new Date());
	 
		return "summernote";
	}
	
}
