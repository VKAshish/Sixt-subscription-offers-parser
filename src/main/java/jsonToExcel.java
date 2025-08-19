import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class jsonToExcel {

    public static void main(String[] args) {
        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Vehicle Offers");

            Row headerRow = sheet.createRow(0);
            String[] headers = {"OwnersOfferKey", "URL", "Maker", "Model", "Term",
                "minContractDurationInMonths", "Version", "transmissionId", "acrissCode", "powerType", "Rate", "StartingFee",
                "AvailableDeliveryLocations", "homeDeliveryFee", "ImageNames", "TypeCategory",
                "DoorCount"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                CellStyle headerStyle = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                headerStyle.setFont(font);
                cell.setCellStyle(headerStyle);
            }

            String jsonContent;
            try {
                java.io.InputStream is = jsonToExcel.class.getClassLoader()
                    .getResourceAsStream("data.json");
                if (is != null) {
                    jsonContent = new String(is.readAllBytes());
                    is.close();
                } else {
                    jsonContent = new String(Files.readAllBytes(Paths.get("data.json")));
                }
            } catch (Exception e) {
                jsonContent = new String(Files.readAllBytes(Paths.get("data.json")));
            }

            Pattern subscriptionPattern = Pattern.compile("\"subscriptionOffers\"\\s*:\\s*\\[",
                Pattern.DOTALL);
            Matcher subscriptionMatcher = subscriptionPattern.matcher(jsonContent);

            int offerCount = 0;
            int rowNum = 1;

            if (subscriptionMatcher.find()) {
                int startPos = subscriptionMatcher.end() - 1;
                int endPos = findMatchingBracket(jsonContent, startPos);

                if (endPos != -1) {
                    String subscriptionSection = jsonContent.substring(startPos + 1, endPos);

                    String[] subscriptionOffers = splitSubscriptionOffers(subscriptionSection);

                    for (String subscriptionOffer : subscriptionOffers) {
                        if (subscriptionOffer.trim().isEmpty()) {
                            continue;
                        }

                        String termLabel = extractValue(subscriptionOffer, "termLabel");

                        Pattern offerPattern = Pattern.compile("\"offers\"\\s*:\\s*\\[",
                            Pattern.DOTALL);
                        Matcher offerMatcher = offerPattern.matcher(subscriptionOffer);

                        while (offerMatcher.find()) {
                            int offerStartPos = offerMatcher.end() - 1;
                            int offerEndPos = findMatchingBracket(subscriptionOffer, offerStartPos);

                            if (offerEndPos != -1) {
                                String offersSection = subscriptionOffer.substring(
                                    offerStartPos + 1, offerEndPos);

                                String[] offers = splitOffers(offersSection);

                                for (String offer : offers) {
                                    if (offer.trim().isEmpty()) {
                                        continue;
                                    }

                                    offerCount++;

                                    String shortSubline = extractValue(offer, "shortSubline");
                                    String acrissCode = extractValue(offer, "acrissCode");
                                    String name = extractValue(offer, "name");
                                    String rate = extractPriceAmount(offer, "monthlyPrice");
                                    String startingFee = extractPriceAmount(offer,
                                        "totalStartingFee");
                                    String imageNames = extractSideImage(offer);
                                    if (imageNames != null && imageNames.startsWith("/")
                                        && !imageNames.equals("Not found")) {
                                        imageNames = imageNames.substring(1);
                                    }
                                    String typeCategory = extractValue(offer, "bodyStyle");
                                    String doorCount = extractNumericValue(offer, "doors");

                                    String minContractDurationInMonths = convertTermToMonths(
                                        termLabel);

                                    String transmissionId = "";
                                    if (acrissCode != null && !acrissCode.equals("Not found")
                                        && acrissCode.length() >= 3) {
                                        transmissionId = String.valueOf(acrissCode.charAt(2));
                                    }

                                    String electric = extractBooleanValue(offer, "electric");
                                    String powerType = "ICE";
                                    if ("true".equals(electric)) {
                                        powerType = "EV";
                                    }

                                    String maker = "";
                                    String model = "";
                                    if (name != null && !name.equals("Not found")) {
                                        String[] parts = name.split("\\s+", 2);
                                        if (parts.length >= 1) {
                                            maker = parts[0];
                                        }
                                        if (parts.length >= 2) {
                                            model = parts[1];
                                        }
                                    }

                                    Row row = sheet.createRow(rowNum++);
                                    row.createCell(0).setCellValue("OwnersOfferKey " + offerCount);
                                    row.createCell(1).setCellValue(
                                        "https://www.sixt.de/plus/offerlist/?acrisscode="
                                            + acrissCode);
                                    row.createCell(2).setCellValue(maker);
                                    row.createCell(3).setCellValue(model);
                                    row.createCell(4).setCellValue(termLabel);
                                    row.createCell(5).setCellValue(minContractDurationInMonths);
                                    row.createCell(6).setCellValue(shortSubline);
                                    row.createCell(7).setCellValue(transmissionId);
                                    row.createCell(8).setCellValue(acrissCode);
                                    row.createCell(9).setCellValue(powerType);
                                    row.createCell(10).setCellValue(rate);
                                    row.createCell(11).setCellValue(startingFee);
                                    row.createCell(12).setCellValue("An vielen SIXT Stationen");
                                    row.createCell(13).setCellValue("199€");
                                    row.createCell(14)
                                        .setCellValue("https://www.sixt.com" + imageNames);
                                    row.createCell(15).setCellValue(typeCategory);
                                    row.createCell(16).setCellValue(doorCount);
                                }
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            FileOutputStream fileOut = new FileOutputStream("vehicle_offers.xlsx");
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();

            System.out.println("Excel file created successfully: vehicle_offers.xlsx");
            System.out.println("Total offers processed: " + offerCount);

        } catch (IOException e) {
            System.err.println("Error reading file or writing Excel: " + e.getMessage());
        } catch (Exception e) {
            System.err.println("Error parsing JSON: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static int findMatchingBracket(String json, int startPos) {
        int bracketCount = 1;
        boolean inString = false;
        boolean escaped = false;

        for (int i = startPos + 1; i < json.length(); i++) {
            char c = json.charAt(i);

            if (escaped) {
                escaped = false;
                continue;
            }

            if (c == '\\') {
                escaped = true;
                continue;
            }

            if (c == '"' && !escaped) {
                inString = !inString;
                continue;
            }

            if (!inString) {
                if (c == '[' || c == '{') {
                    bracketCount++;
                } else if (c == ']' || c == '}') {
                    bracketCount--;
                    if (bracketCount == 0) {
                        return i;
                    }
                }
            }
        }

        return -1;
    }

    private static String[] splitSubscriptionOffers(String subscriptionSection) {
        java.util.List<String> subscriptionOffers = new java.util.ArrayList<>();
        int braceCount = 0;
        int start = 0;
        boolean inString = false;
        boolean escaped = false;

        for (int i = 0; i < subscriptionSection.length(); i++) {
            char c = subscriptionSection.charAt(i);

            if (escaped) {
                escaped = false;
                continue;
            }

            if (c == '\\') {
                escaped = true;
                continue;
            }

            if (c == '"' && !escaped) {
                inString = !inString;
                continue;
            }

            if (!inString) {
                if (c == '{') {
                    if (braceCount == 0) {
                        start = i;
                    }
                    braceCount++;
                } else if (c == '}') {
                    braceCount--;
                    if (braceCount == 0) {
                        subscriptionOffers.add(subscriptionSection.substring(start, i + 1));
                    }
                }
            }
        }

        return subscriptionOffers.toArray(new String[0]);
    }

    private static String[] splitOffers(String offersSection) {
        java.util.List<String> offers = new java.util.ArrayList<>();
        int braceCount = 0;
        int start = 0;
        boolean inString = false;
        boolean escaped = false;

        for (int i = 0; i < offersSection.length(); i++) {
            char c = offersSection.charAt(i);

            if (escaped) {
                escaped = false;
                continue;
            }

            if (c == '\\') {
                escaped = true;
                continue;
            }

            if (c == '"' && !escaped) {
                inString = !inString;
                continue;
            }

            if (!inString) {
                if (c == '{') {
                    if (braceCount == 0) {
                        start = i;
                    }
                    braceCount++;
                } else if (c == '}') {
                    braceCount--;
                    if (braceCount == 0) {
                        offers.add(offersSection.substring(start, i + 1));
                    }
                }
            }
        }

        return offers.toArray(new String[0]);
    }

    private static String extractValue(String json, String key) {
        Pattern pattern = Pattern.compile("\"" + key + "\"\\s*:\\s*\"([^\"]*?)\"");
        Matcher matcher = pattern.matcher(json);
        if (matcher.find()) {
            return matcher.group(1);
        }
        return "Not found";
    }

    private static String extractNumericValue(String json, String key) {
        Pattern pattern = Pattern.compile("\"" + key + "\"\\s*:\\s*([0-9]+\\.?[0-9]*)");
        Matcher matcher = pattern.matcher(json);
        if (matcher.find()) {
            return matcher.group(1);
        }
        return "Not found";
    }

    private static String extractPriceAmount(String json, String key) {
        Pattern pattern = Pattern.compile("\"" + key
                + "\"\\s*:\\s*\\{[^}]*\"amount\"\\s*:\\s*\\{[^}]*\"value\"\\s*:\\s*([0-9]+\\.?[0-9]*)",
            Pattern.DOTALL);
        Matcher matcher = pattern.matcher(json);
        if (matcher.find()) {
            return matcher.group(1);
        }
        return "Not found";
    }

    private static String extractSideImage(String json) {
        Pattern pattern = Pattern.compile(
            "\"sideImages\"\\s*:\\s*\\{[^}]*\"large\"\\s*:\\s*\"([^\"]*?)\"", Pattern.DOTALL);
        Matcher matcher = pattern.matcher(json);
        if (matcher.find()) {
            return matcher.group(1);
        }
        return "Not found";
    }

    private static String extractBooleanValue(String json, String key) {
        Pattern pattern = Pattern.compile("\"" + key + "\"\\s*:\\s*(true|false)");
        Matcher matcher = pattern.matcher(json);
        if (matcher.find()) {
            return matcher.group(1);
        }
        return "false";
    }

    private static String convertTermToMonths(String termLabel) {
        if (termLabel == null || termLabel.equals("Not found")) {
            return "Not found";
        }

        if (termLabel.contains("1 Monat")) {
            return "1";
        } else if (termLabel.contains("6 Monate")) {
            return "6";
        } else if (termLabel.contains("12 Monate")) {
            return "12";
        }

        return "Not found";
    }
}
