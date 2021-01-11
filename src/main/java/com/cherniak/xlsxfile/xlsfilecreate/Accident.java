package com.cherniak.xlsxfile.xlsfilecreate;

import java.time.Instant;
import java.time.LocalDate;
import java.util.Date;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;


@Data
@NoArgsConstructor
@AllArgsConstructor
public class Accident {


  private LocalDate createDate;
  private String extId;
  private String policyNum;
  private String skName;
  private LocalDate accidentDate;
  private String driver;
  private String docType;
  private String docNum;
  private String vin;
  private String regNum;
  private String bodyNum;
  private String iss;
}
