package com.example.customermanagement.Controller;

import com.example.customermanagement.Exception.ResourceNotFoundException;
import com.example.customermanagement.model.Customer;
import com.example.customermanagement.repository.CustomerRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/v1")
public class CustomerController {

    @Autowired
    public CustomerRepository customerRepository;

    @GetMapping("/customers")
    public List<Customer> getAllCustomers(){
        return  customerRepository.findAll();
    }

    @GetMapping("/customers/{id}")
    public ResponseEntity<Customer> getCustomerById(@PathVariable(value = "id") Long customerId)
            throws ResourceNotFoundException {

        Customer customer = customerRepository.findById(customerId)
                .orElseThrow(() -> new ResourceNotFoundException("Customer not found for this id ::" +customerId));
        return ResponseEntity.ok().body(customer);

    }

    @PostMapping("/customers")
    public Customer createCustomer(@Validated @RequestBody Customer customer){
        return customerRepository.save(customer);
    }

    @PutMapping("customers/{id}")
    public ResponseEntity<Customer> updateCustomer(@PathVariable(value = "id") Long customerId,
                                                   @Validated @RequestBody Customer customerDetails) throws ResourceNotFoundException{
        Customer customer = customerRepository.findById(customerId)
                .orElseThrow(() -> new ResourceNotFoundException("Customer Not found this id :: "+customerId));

        customer.setName(customerDetails.getName());
        customer.setDateOfBirth(customerDetails.getDateOfBirth());
        customer.setNicNumber(customerDetails.getNicNumber());
        customer.setPhoneNumber(customerDetails.getPhoneNumber());
        customer.setAddressLine1(customerDetails.getAddressLine1());
        customer.setAddressLine2(customerDetails.getAddressLine2());
        customer.setCity(customerDetails.getCity());
        customer.setCountry(customerDetails.getCountry());
        final Customer updateCustomer = customerRepository.save(customer);
        return  ResponseEntity.ok(updateCustomer);

    }

    @DeleteMapping("/customers/{id}")
    public Map<String,Boolean> deleteCustomer(@PathVariable(value = "id") Long customerId)
            throws ResourceNotFoundException{
        Customer customer  = customerRepository.findById(customerId)
                .orElseThrow(() ->new ResourceNotFoundException("Customer not found for this id ::"+customerId));

        customerRepository.delete(customer);
        Map<String,Boolean> response = new HashMap<>();
        response.put("deleted",Boolean.TRUE);
        return response;

    }
}
