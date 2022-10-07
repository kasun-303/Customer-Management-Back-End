package com.example.customermanagement.model;

import com.fasterxml.jackson.annotation.JsonBackReference;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;

import javax.persistence.*;
import javax.xml.stream.events.Comment;

@Data
@Getter
@Setter
@Entity
@Table(name = "customers")
public class Customer {

    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private long id;
    private String name;
    private String dateOfBirth;
    private String nicNumber;
    private String phoneNumber;

//    @ManyToOne
//    @JoinColumn(referencedColumnName = "id")
//    @JsonBackReference
//    private Customer familyMember;

    private String addressLine1;
    private String addressLine2;
    private String city;
    private String country;

}