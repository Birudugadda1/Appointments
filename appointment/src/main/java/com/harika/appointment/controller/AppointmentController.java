package com.harika.appointment.controller;


import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.harika.appointment.dao.AppointmentRepository;
import com.harika.appointment.model.Appointment;


@RestController
@RequestMapping("/appointments")
public class AppointmentController {
	@Autowired
	private AppointmentRepository appointmentRepository;
	
	
	@GetMapping
	public List<Appointment> getAppointments() {
  
    	return appointmentRepository.findAll();
    }
	
	
	@PostMapping
	public void create(@RequestBody Appointment a) {
		appointmentRepository.save(a);
    }
	
	@GetMapping("/{appointmentid}")
	public Appointment getProduct(@PathVariable("appointmentid") long AppoitmentId) {
		return appointmentRepository.getOne(AppoitmentId);
	}
		

}
