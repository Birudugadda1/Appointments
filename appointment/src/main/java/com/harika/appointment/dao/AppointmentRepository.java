package com.harika.appointment.dao;

import org.springframework.data.jpa.repository.JpaRepository;

import com.harika.appointment.model.Appointment;

public interface AppointmentRepository extends JpaRepository<Appointment, Long>{

}
