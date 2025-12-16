package com.paymentrecord.controller;

import com.paymentrecord.dto.PaymentRequest;
import com.paymentrecord.service.GoogleSheetService;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/payment")
public class PaymentController {

    private final GoogleSheetService service;

    public PaymentController(GoogleSheetService service){
        this.service = service;
    }

    @PostMapping("/save")
    public String save(@RequestBody PaymentRequest request) throws Exception {
        service.savePayment(request);
        return "Payment Stored Successfully";
    }

}
