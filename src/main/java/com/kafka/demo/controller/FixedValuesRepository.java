package com.kafka.demo.controller;

import org.springframework.data.repository.CrudRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface FixedValuesRepository extends CrudRepository<FixedValues, Integer> {}