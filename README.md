# Sixt Subscription Offers Parser

A Java application that parses Sixt subscription offers from JSON data and exports them to Excel format.

## Features

- Parses JSON data containing vehicle subscription offers
- Extracts vehicle information (maker, model, version, etc.)
- Processes multiple subscription terms (1 month, 6 months, 12 months)
- Exports data to Excel (.xlsx) format with proper formatting
- Handles image URLs and pricing information

## Output Columns

The generated Excel file contains the following columns:

1. **OwnersOfferKey** - Unique identifier for each offer
2. **URL** - Sixt URL with acriss code
3. **Maker** - Vehicle manufacturer
4. **Model** - Vehicle model
5. **Term** - Subscription duration ("1 Monat", "6 Monate", "12 Monate")
6. **minContractDurationInMonths** - Numeric contract duration (1, 6, or 12)
7. **Version** - Short subline description
8. **transmissionId** - 3rd letter of acriss code
9. **Rate** - Monthly price amount
10. **StartingFee** - Total starting fee amount
11. **AvailableDeliveryLocations** - Delivery locations
12. **homeDeliveryFee** - Home delivery fee (199€)
13. **ImageNames** - Vehicle image URLs
14. **TypeCategory** - Body style
15. **DoorCount** - Number of doors

## Requirements

- Java 11 or higher
- Gradle

## Usage

1. Place your JSON data file as `src/main/resources/data.json`
2. Run the application:
   ```bash
   gradle run
   ```
3. The Excel file `vehicle_offers.xlsx` will be generated in the project root

## Build

```bash
gradle build
```

## Dependencies

- Apache POI 5.2.4 - For Excel file generation

## Data Structure

The application expects JSON data with a `subscriptionOffers` array containing vehicle offers organized by subscription terms. 