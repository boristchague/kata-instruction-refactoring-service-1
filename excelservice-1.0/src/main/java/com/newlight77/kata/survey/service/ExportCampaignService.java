package com.newlight77.kata.survey.service;

import com.newlight77.kata.survey.model.AddressStatus;
import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import com.newlight77.kata.survey.client.CampaignClient;
import com.newlight77.kata.survey.utils.SaveToExcelFileService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

@Component
public class ExportCampaignService {

  private final CampaignClient campaignWebService;
  private final MailService mailService;
  private final SaveToExcelFileService saveToExcelFileService;

  public ExportCampaignService(final CampaignClient campaignWebService, final MailService mailService, final SaveToExcelFileService saveToExcelFileService) {
    this.campaignWebService = campaignWebService;
    this.mailService = mailService;
    this.saveToExcelFileService = saveToExcelFileService;
  }

  public void creerSurvey(final Survey survey) {
    campaignWebService.createSurvey(survey);
  }

  public Survey getSurvey(final String id) {
    return campaignWebService.getSurvey(id);
  }

  public void createCampaign(final Campaign campaign) {
    campaignWebService.createCampaign(campaign);
  }

  public Campaign getCampaign(final String id) {
    return campaignWebService.getCampaign(id);
  }

  public void sendResults(final Campaign campaign, final Survey survey) {

    final Workbook workbook = new XSSFWorkbook();

    final Sheet sheet = workbook.createSheet("Survey");

    saveToExcelFileService.setColumnWidths(sheet);

    // 1ere ligne =  l'entête
    final Row header = sheet.createRow(0);

    //create headerstyle
    final CellStyle headerStyle = saveToExcelFileService.createHeaderStyle(workbook);

    //heading of the survey
    final Cell headerCell = header.createCell(0);
    headerCell.setCellValue("Survey");
    headerCell.setCellStyle(headerStyle);

    //set style for each celle for survey data
    final CellStyle titleStyle = saveToExcelFileService.createTitleStyle(workbook);

    final CellStyle style = workbook.createCellStyle();
    style.setWrapText(true);

    // section client
    saveToExcelFileService.createClientSection(sheet, survey, titleStyle, style);

    saveToExcelFileService.createSurveyInfo(sheet, campaign, style);

    saveToExcelFileService.createSurveyHeader(sheet, style);

    saveToExcelFileService.fillSurveyData(sheet, campaign, style);

    saveToExcelFileService.writeFileAndSend(survey, workbook);

  }

}


