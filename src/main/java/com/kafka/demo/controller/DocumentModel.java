package com.kafka.demo.controller;

import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.annotation.JsonProperty;

public class DocumentModel {
	
	/**
	 * 
	 */
	@JsonProperty("document")
	private String document;
	@JsonProperty("managers")
	private List<Integer> managers ;
	@JsonProperty("date")
	private String date ;
	
	public String getDocument() {
		return document;
	}

	public void setDocument(String document) {
		this.document = document;
	}

	public List<Integer> getManagers() {
		if(managers == null)
			managers = new ArrayList<>();
		return managers;
	}

	public void setManagers(List<Integer> managers) {
		this.managers = managers;
	}

	public String getDate() {
		return date;
	}

	public void setDate(String date) {
		this.date = date;
	}

	
	
}
