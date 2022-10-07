package com.example.customermanagement.service;

import java.io.IOException;
import java.util.List;

import com.example.customermanagement.helper.ExcelHelper;
import com.example.customermanagement.model.Customer;
import com.example.customermanagement.repository.CustomerRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;


@Service
public class ExcelService {

    @Autowired
    CustomerRepository repository;

    public void save(MultipartFile file) {
        try {
            List<Customer> customers = ExcelHelper.excelToCustomers(file.getInputStream());
            repository.saveAll(customers);
        } catch (IOException e) {
            throw new RuntimeException("fail to store excel data: " + e.getMessage());
        }
    }

    public List<Customer> getAllCustomers() {
        return repository.findAll();
    }
}
