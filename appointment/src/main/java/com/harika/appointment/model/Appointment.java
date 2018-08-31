package com.harika.appointment.model;

import java.util.Date;

public class Appointment {
	private int AppoitmentId;
	private Date date;
	private String description;
	public int getAppoitmentId() {
		return AppoitmentId;
	}
	public void setAppoitmentId(int appoitmentId) {
		AppoitmentId = appoitmentId;
	}
	public Date getDate() {
		return date;
	}
	public void setDate(Date date) {
		this.date = date;
	}
	public String getDescription() {
		return description;
	}
	public void setDescription(String description) {
		this.description = description;
	}
}
