package com.example.execsheetfileration.ResponseDTO;

import lombok.*;

@Getter
@Setter
@Data
@AllArgsConstructor
@NoArgsConstructor
public class DataResponse {

    private int countEMERALD;
    private int countSPEAR;
    private int countCUSHION;
    private int countOVAL;
    private int countPRINCESS;
    private int countSMARQUISE;
    private int totalType;
    private double totalWeightOfEmerald;
    private double totalWeightOfPear;
    private double totalWeightOfCushion;
    private double totalWeightOfOval;
    private double totalWeightOfPrincess;
    private double totalWeightOfMarquise;
    private double totalkachuWeight;

    private double totalPolishedWeightOfEmerald;
    private double totalPolishedWeightOfPear;
    private double totalPolishedWeightOfCushion;
    private double totalPolishedWeightOfOval;
    private double totalPolishedWeightOfPrincess;
    private double totalPolishedWeightOfMarquise;
    private double totalPolishedWeight;

}
