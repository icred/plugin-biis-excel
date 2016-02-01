package eu.icred.external.plugin.biis.xssf.read;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Currency;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.SortedMap;
import java.util.regex.Pattern;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.LocalDate;
import org.joda.time.LocalDateTime;
import org.joda.time.Period;

import eu.icred.model.datatype.Amount;
import eu.icred.model.datatype.Area;
import eu.icred.model.datatype.enumeration.AreaMeasurement;
import eu.icred.model.datatype.enumeration.AreaType;
import eu.icred.model.datatype.enumeration.ConstructionPhase;
import eu.icred.model.datatype.enumeration.Country;
import eu.icred.model.datatype.enumeration.InteriorQuality;
import eu.icred.model.datatype.enumeration.ObjectCondition;
import eu.icred.model.datatype.enumeration.OwnershipType;
import eu.icred.model.datatype.enumeration.RetailLocationType;
import eu.icred.model.datatype.enumeration.Subset;
import eu.icred.model.datatype.enumeration.UseType;
import eu.icred.model.datatype.enumeration.ValuationType1;
import eu.icred.model.datatype.enumeration.ValuationType2;
import eu.icred.model.node.Container;
import eu.icred.model.node.Data;
import eu.icred.model.node.Meta;
import eu.icred.model.node.entity.Property;
import eu.icred.model.node.entity.Valuation;
import eu.icred.model.node.group.Address;
import eu.icred.plugin.PluginComponent;
import eu.icred.plugin.worker.WorkerConfiguration;
import eu.icred.plugin.worker.input.IImportWorker;
import eu.icred.plugin.worker.input.ImportWorkerConfiguration;
import eu.icred.validator.subset_5_7.ValuationValidator;

public class Reader implements IImportWorker {
    private static Logger logger = Logger.getLogger(Reader.class);

    public static final Subset[] SUPPORTED_SUBSETS = { Subset.S5_7 };
    private static final String PARAMETER_NAME_STREAM = "biis-file";
    private static final String PARAMETER_NAME_SHEET_IDX = "sheet-number";
    private static final String PARAMETER_NAME_SHEET_NAME = "sheet-name";

    private Container container = null;
    private InputStream xmlStream = null;

    @Override
    public List<Subset> getSupportedSubsets() {
        return Arrays.asList(SUPPORTED_SUBSETS);
    }

    @Override
    public void load(WorkerConfiguration config) {
        throw new RuntimeException("not allowed");
    }

    @Override
    public void unload() {
        try {
            xmlStream.close();
        } catch (Throwable t) {
        }
        xmlStream = null;
    }

    @Override
    public void load(ImportWorkerConfiguration config) {
        xmlStream = config.getStreams().get(PARAMETER_NAME_STREAM);

        String sheetName = config.getStrings().get(PARAMETER_NAME_SHEET_NAME);
        Integer sheetIdx = config.getIntegers().get(PARAMETER_NAME_SHEET_IDX);

        XSSFWorkbook workbook = null;
        try {

            workbook = new XSSFWorkbook(xmlStream);
            XSSFSheet sheet = null;

            if (sheetIdx != null) {
                sheet = workbook.getSheetAt(sheetIdx - 1);
            } else {
                if (sheetName != null) {
                    sheet = workbook.getSheet(sheetName);
                }
            }

            container = new Container();

            Meta meta = container.getMeta();
            meta.setCreated(LocalDateTime.now());
            meta.setCreator("icred with biis-excel plugin");
            meta.setFormat("XML");
            meta.setVersion("1-0.6.2");

            Data data = container.getMaindata();
            Map<String, Property> properties = data.getProperties();

            Map<Integer, String> headers = new HashMap<Integer, String>();
            int rowsCount = sheet.getLastRowNum();

            for (int rowIndex = 0; rowIndex <= rowsCount; rowIndex++) {

                Property prop = null;
                Valuation val = null;
                Address valAddress = null;

                Currency mainCurrency = null;
                AreaMeasurement mainAreaMeasurement = null;

                try {
                    Row row = sheet.getRow(rowIndex);

                    if (row.getCell(0) == null) {
                        continue;
                    }

                    if (rowIndex != 0) {
                        prop = new Property();
                        val = new Valuation();
                        valAddress = new Address();
                        val.setAddress(valAddress);
                    }

                    int colCounts = row.getLastCellNum();
                    for (int columnIndex = 0; columnIndex < colCounts; columnIndex++) {

                        Cell cell = row.getCell(columnIndex);
                        if (cell != null) {
                            if (rowIndex == 0) {
                                headers.put(columnIndex, getCellStringValue(cell));
                            } else {
                                String biisKeyName = headers.get(columnIndex);
                                try {
                                    if (biisKeyName.equals("Date")) {
                                        // ignore - see field DateOfAppraisal

                                    } else if (biisKeyName.equals("CompletionDate")) {
                                        val.setValuationDate(biis2gif_Date(cell));

                                    } else if (biisKeyName.equals("DataSupplier")) {
                                        val.setLabel(getCellStringValue(cell));
                                        val.setExpertName(getCellStringValue(cell));

                                    } else if (biisKeyName.equals("TypeOfDataSupplier")) {
                                        // ignore

                                    } else if (biisKeyName.equals("ArealUnit")) {
                                        mainAreaMeasurement = biis2gif_AreaMeasureMent(getCellStringValue(cell));

                                    } else if (biisKeyName.equals("AddressType_Street")) {
                                        valAddress.setStreet(getCellStringValue(cell));

                                    } else if (biisKeyName.equals("AddressType_PostCode")) {
                                        valAddress.setZip(getCellStringValue(cell));

                                    } else if (biisKeyName.equals("AddressType_Town")) {
                                        valAddress.setCity(getCellStringValue(cell));

                                    } else if (biisKeyName.equals("AddressType_ISOCountryCodeType_Country")) {
                                        String country = getCellStringValue(cell);
                                        if (country != null && country.length() > 0)
                                            valAddress.setCountry(Country.valueOf(country));

                                    } else if (biisKeyName.equals("AddressType_Text")) {
                                        String label = getCellStringValue(cell);
                                        if (prop.getLabel() == null) {
                                            prop.setLabel(label);
                                        }
                                        valAddress.setLabel(label);

                                    } else if (biisKeyName.equals("Owner")) {
                                        val.setOwner(getCellStringValue(cell));

                                    } else if (biisKeyName.equals("ObjNoOwner")) {
                                        String propId = getCellStringValue(cell);

                                        if (properties.get(propId) == null) {
                                            properties.put(propId, prop);

                                            prop.setObjectIdSender(propId);
                                            prop.setObjectIdReceiver(propId);

                                        } else {
                                            prop = properties.get(propId);
                                        }

                                    } else if (biisKeyName.equals("ObjKoWGS84Longitude")) {
                                        valAddress.setLongitude(biis2gif_Double(cell));

                                    } else if (biisKeyName.equals("ObjKoWGS84Latitude")) {
                                        valAddress.setLatitude(biis2gif_Double(cell));

                                    } else if (biisKeyName.equals("RebaseType1")) {
                                        val.setValuationType1(biis2gif_ValuationType1(getCellStringValue(cell)));

                                    } else if (biisKeyName.equals("RebaseType2")) {
                                        val.setValuationType2(biis2gif_ValuationType2(getCellStringValue(cell)));

                                    } else if (biisKeyName.equals("RebaseObjAdditionalInformation")) {
                                        val.setNote(getCellStringValue(cell));

                                    } else if (biisKeyName.equals("DateOfAppraisal")) {
                                        val.setValidFrom(biis2gif_Date(cell));
                                        setObjectId(prop, val);

                                    } else if (biisKeyName.equals("QualityDateOfAppraisal")) {
                                        // ignore

                                    } else if (biisKeyName.equals("Currency")) {
                                        mainCurrency = Currency.getInstance(getCellStringValue(cell));
                                        val.setCurrency(mainCurrency);

                                    } else if (biisKeyName.equals("ExchangeRate1EUR")) {
                                        val.setExchangeRateToEUR(biis2gif_Double(cell));

                                    } else if (biisKeyName.equals("DateExchangeRate")) {
                                        val.setExchangeRateDate(biis2gif_Date(cell));

                                    } else if (biisKeyName.equals("MainTypeOfUse")) {
                                        val.setUseTypePrimary(biis2gif_UseType(getCellStringValue(cell)));

                                    } else if (biisKeyName.equals("ShareMainTypeOfUse")) {
                                        val.setUseTypePrimaryShare(biis2gif_Double(cell));

                                    } else if (biisKeyName.equals("AncillaryTypeOfUse")) {
                                        val.setUseTypeSecondary(biis2gif_UseType(getCellStringValue(cell)));

                                    } else if (biisKeyName.equals("ShareAncillaryTypeOfUse")) {
                                        val.setUseTypeSecondaryShare(biis2gif_Double(cell));

                                    } else if (biisKeyName.equals("TypeOfOwnership")) {
                                        val.setOwnershipType(biis2gif_OwnershipType(cell));

                                    } else if (biisKeyName.equals("SingleTenant")) {
                                        val.setSingleTenant(biis2gif_Boolean(cell));

                                    } else if (biisKeyName.equals("PurchasePrice")) {
                                        val.setPurchaseNetPrice(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("DateOfPurchase")) {
                                        val.setPurchaseDate(biis2gif_Date(cell));

                                    } else if (biisKeyName.equals("PriceOfSale")) {
                                        val.setSaleNetPrice(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("DateOfSale")) {
                                        val.setSaleDate(biis2gif_Date(cell));

                                    } else if (biisKeyName.equals("LocationQuality")) {
                                        val.setRetailLocation(biis2gif_RetailLocationType(getCellStringValue(cell)));

                                    } else if (biisKeyName.equals("StructuralCondition")) {
                                        val.setCondition(biis2gif_Condition(getCellStringValue(cell)));

                                    } else if (biisKeyName.equals("FitOutQuality")) {
                                        val.setInteriorQuality(biis2gif_InteriorQuality(getCellStringValue(cell)));

                                    } else if (biisKeyName.equals("StateOfCompletion")) {
                                        val.setConstructionPhase(biis2gif_ConstructionPhase(getCellStringValue(cell)));

                                    } else if (biisKeyName.equals("MaintenanceBacklog")) {
                                        val.setMaintenanceBacklog(biis2gif_Boolean(cell));

                                    } else if (biisKeyName.equals("Floors")) {
                                        val.setFloorDescription(getCellStringValue(cell));

                                    } else if (biisKeyName.equals("NormalTotalEconomicLife")) {
                                        Double yearVal = biis2gif_Double(cell);
                                        if (yearVal != null)
                                            val.setNormalTotalEconomicLife(org.joda.time.Period.years(yearVal.intValue()));

                                    } else if (biisKeyName.equals("RemainingEconomicLife")) {
                                        Double yearVal = biis2gif_Double(cell);
                                        if (yearVal != null)
                                            val.setRemainingEconomicLife(org.joda.time.Period.years(yearVal.intValue()));

                                    } else if (biisKeyName.equals("OriginalYearOfConstruction")) {
                                        val.setConstructionDate(biis2gif_Year(cell));

                                    } else if (biisKeyName.equals("CalculatedYearOfConstruction")) {
                                        val.setEconomicConstructionDate(biis2gif_Year(cell));

                                    } else if (biisKeyName.equals("DateOfChangeForRemainingEconomicLife")) {
                                        val.setChangeDateForRemainingEconomicLife(biis2gif_Date(cell));

                                    } else if (biisKeyName.equals("LandSize")) {
                                        val.setPlotArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("FloorToAreaRatio")) {
                                        val.setGfz(biis2gif_Double(cell));

                                    } else if (biisKeyName.equals("SiteCoverageRatio")) {
                                        val.setGrz(biis2gif_Double(cell));

                                    } else if (biisKeyName.equals("GrossFloorSpaceOverground")) {
                                        val.setGrossFloorSpaceOverground(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("GrossFloorSpaceBelowGround")) {
                                        val.setGrossFloorSpaceBelowGround(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("TotalGrossFloorSpace")) {
                                        val.setTotalGrossFloorSpace(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("TotalRentableArea")) {
                                        val.setTotalRentableArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RunningCosts")) {
                                        val.setRunningCosts(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("ManagementCosts")) {
                                        val.setManagementCosts(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("MaintenanceExpenses")) {
                                        val.setMaintenanceExpenses(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentAllowance")) {
                                        val.setRentAllowance(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("OtherOperatingExpenses")) {
                                        val.setOtherOperatingExpenses(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("CapitalizationRate")) {
                                        val.setCapitalizationRate(biis2gif_Double(cell));

                                    } else if (biisKeyName.equals("ValueByIncomeApproachWithoutPremiumsDiscounts")) {
                                        val.setValueByIncomeApproachWithoutPremiumsDiscounts(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("DiscountsPremiums")) {
                                        val.setDiscountsPremiums(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("DeductionForVacancy")) {
                                        val.setDeductionForVacancy(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("DeductionConstructionWorks")) {
                                        val.setDeductionConstructionWorks(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("OthersDiscountsPremiums")) {
                                        val.setOthersDiscountsPremiums(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("ValueByIncomeApproach")) {
                                        val.setValueByIncomeApproach(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("CostApproach")) {
                                        val.setCostApproach(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("LandValue")) {
                                        val.setLandValue(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("MarketValue")) {
                                        val.setFairValue(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("GroundLease")) {
                                        val.setGroundLease(biis2gif_Boolean(cell));

                                    } else if (biisKeyName.equals("RemainingLifeOfGroundLease")) {
                                        val.setRemainingLifeOfGroundLease(biis2gif_Period(cell));
                                        
                                    } else if (biisKeyName.equals("GroundRent")) {
                                        val.setGroundRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("GroundLeaseRemarks")) {
                                        val.setGroundLeaseRemarks(getCellStringValue(cell));

                                    } else if (biisKeyName.equals("RentalSituationOfficeLetArea")) {
                                        val.setRentalSituationOfficeLetArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationOfficeContractualAnnualRent")) {
                                        val.setRentalSituationOfficeContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationOfficeEstimatedAnnualRentForLetArea")) {
                                        val.setRentalSituationOfficeEstimatedAnnualRentForLetArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationOfficeVacantArea")) {
                                        val.setRentalSituationOfficeVacantArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationOfficeEstimatedAnnualRentForVacantArea")) {
                                        val.setRentalSituationOfficeEstimatedAnnualRentForVacantArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationRetailLetArea")) {
                                        val.setRentalSituationRetailLetArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationRetailContractualAnnualRent")) {
                                        val.setRentalSituationRetailContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationRetailEstimatedAnnualRentForLetArea")) {
                                        val.setRentalSituationRetailEstimatedAnnualRentForLetArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationRetailVacantArea")) {
                                        val.setRentalSituationRetailVacantArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationRetailEstimatedAnnualRentForVacantArea")) {
                                        val.setRentalSituationRetailEstimatedAnnualRentForVacantArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationStorageLetArea")) {
                                        val.setRentalSituationStorageLetArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationStorageContractualAnnualRent")) {
                                        val.setRentalSituationStorageContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationStorageEstimatedAnnualRentForLetArea")) {
                                        val.setRentalSituationStorageEstimatedAnnualRentForLetArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationStorageVacantArea")) {
                                        val.setRentalSituationStorageVacantArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationStorageEstimatedAnnualRentForVacantArea")) {
                                        val.setRentalSituationStorageEstimatedAnnualRentForVacantArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationArchiveLetArea")) {
                                        val.setRentalSituationArchiveLetArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationArchiveContractualAnnualRent")) {
                                        val.setRentalSituationArchiveContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationArchiveEstimatedAnnualRentForLetArea")) {
                                        val.setRentalSituationArchiveEstimatedAnnualRentForLetArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationArchiveVacantArea")) {
                                        val.setRentalSituationArchiveVacantArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationArchiveEstimatedAnnualRentForVacantArea")) {
                                        val.setRentalSituationArchiveEstimatedAnnualRentForVacantArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationGastroLetArea")) {
                                        val.setRentalSituationGastroLetArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationGastroContractualAnnualRent")) {
                                        val.setRentalSituationGastroContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationGastroEstimatedAnnualRentForLetArea")) {
                                        val.setRentalSituationGastroEstimatedAnnualRentForLetArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationGastroVacantArea")) {
                                        val.setRentalSituationGastroVacantArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationGastroEstimatedAnnualRentForVacantArea")) {
                                        val.setRentalSituationGastroEstimatedAnnualRentForVacantArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationResidentialLetArea")) {
                                        val.setRentalSituationResidentialLetArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationResidentialContractualAnnualRent")) {
                                        val.setRentalSituationResidentialContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationResidentialEstimatedAnnualRentForLetArea")) {
                                        val.setRentalSituationResidentialEstimatedAnnualRentForLetArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationResidentialVacantArea")) {
                                        val.setRentalSituationResidentialVacantArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationResidentialEstimatedAnnualRentForVacantArea")) {
                                        val.setRentalSituationResidentialEstimatedAnnualRentForVacantArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationHotelLetArea")) {
                                        val.setRentalSituationHotelLetArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationHotelContractualAnnualRent")) {
                                        val.setRentalSituationHotelContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationHotelEstimatedAnnualRentForLetArea")) {
                                        val.setRentalSituationHotelEstimatedAnnualRentForLetArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationHotelVacantArea")) {
                                        val.setRentalSituationHotelVacantArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationHotelEstimatedAnnualRentForVacantArea")) {
                                        val.setRentalSituationHotelEstimatedAnnualRentForVacantArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationLeisureLetArea")) {
                                        val.setRentalSituationLeisureLetArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationLeisureContractualAnnualRent")) {
                                        val.setRentalSituationLeisureContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationLeisureEstimatedAnnualRentForLetArea")) {
                                        val.setRentalSituationLeisureEstimatedAnnualRentForLetArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationLeisureVacantArea")) {
                                        val.setRentalSituationLeisureVacantArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationLeisureEstimatedAnnualRentForVacantArea")) {
                                        val.setRentalSituationLeisureEstimatedAnnualRentForVacantArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationIndoorparkingLetNumbers")) {
                                        Double number = biis2gif_Double(cell);
                                        if (number != null)
                                            val.setRentalSituationIndoorparkingLetNumbers(number.intValue());

                                    } else if (biisKeyName.equals("RentalSituationIndoorparkingContractualAnnualRent")) {
                                        val.setRentalSituationIndoorparkingContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationIndoorparkingEstimatedAnnualRentForLetNumbers")) {
                                        val.setRentalSituationIndoorparkingEstimatedAnnualRentForLetNumbers(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationIndoorparkingVacantNumbers")) {
                                        Double number = biis2gif_Double(cell);
                                        if (number != null)
                                            val.setRentalSituationIndoorparkingVacantNumbers(number.intValue());

                                    } else if (biisKeyName.equals("RentalSituationIndoorparkingEstimatedAnnualRentForVacantNumbers")) {
                                        val.setRentalSituationIndoorparkingEstimatedAnnualRentForVacantNumbers(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationOutsideparkingLetNumbers")) {
                                        Double number = biis2gif_Double(cell);
                                        if (number != null)
                                            val.setRentalSituationOutsideparkingLetNumbers(number.intValue());

                                    } else if (biisKeyName.equals("RentalSituationOutsideparkingContractualAnnualRent")) {
                                        val.setRentalSituationOutsideparkingContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationOutsideparkingEstimatedAnnualRentForLetNumbers")) {
                                        val.setRentalSituationOutsideparkingEstimatedAnnualRentForLetNumbers(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationOutsideparkingVacantNumbers")) {
                                        Double number = biis2gif_Double(cell);
                                        if (number != null)
                                            val.setRentalSituationOutsideparkingVacantNumbers(number.intValue());

                                    } else if (biisKeyName.equals("RentalSituationOutsideparkingEstimatedAnnualRentForVacantNumbers")) {
                                        val.setRentalSituationOutsideparkingEstimatedAnnualRentForVacantNumbers(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscArea1LetArea")) {
                                        val.setRentalSituationMiscArea1LetArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationMiscArea1ContractualAnnualRent")) {
                                        val.setRentalSituationMiscArea1ContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscArea1EstimatedAnnualRentForLetArea")) {
                                        val.setRentalSituationMiscArea1EstimatedAnnualRentForLetArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscArea1VacantArea")) {
                                        val.setRentalSituationMiscArea1VacantArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationMiscArea1EstimatedAnnualRentForVacantArea")) {
                                        val.setRentalSituationMiscArea1EstimatedAnnualRentForVacantArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscArea2LetArea")) {
                                        val.setRentalSituationMiscArea2LetArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationMiscArea2ContractualAnnualRent")) {
                                        val.setRentalSituationMiscArea2ContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscArea2EstimatedAnnualRentForLetArea")) {
                                        val.setRentalSituationMiscArea2EstimatedAnnualRentForLetArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscArea2VacantArea")) {
                                        val.setRentalSituationMiscArea2VacantArea(biis2gif_Area(cell, mainAreaMeasurement));

                                    } else if (biisKeyName.equals("RentalSituationMiscArea2EstimatedAnnualRentForVacantArea")) {
                                        val.setRentalSituationMiscArea2EstimatedAnnualRentForVacantArea(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscnumbers1LetNumbers")) {
                                        Double number = biis2gif_Double(cell);
                                        if (number != null)
                                            val.setRentalSituationMiscnumbers1LetNumbers(number.intValue());

                                    } else if (biisKeyName.equals("RentalSituationMiscnumbers1ContractualAnnualRent")) {
                                        val.setRentalSituationMiscnumbers1ContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscnumbers1EstimatedAnnualRentForLetNumbers")) {
                                        val.setRentalSituationMiscnumbers1EstimatedAnnualRentForLetNumbers(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscnumbers1VacantNumbers")) {
                                        Double number = biis2gif_Double(cell);
                                        if (number != null)
                                            val.setRentalSituationMiscnumbers1VacantNumbers(number.intValue());

                                    } else if (biisKeyName.equals("RentalSituationMiscnumbers1EstimatedAnnualRentForVacantNumbers")) {
                                        val.setRentalSituationMiscnumbers1EstimatedAnnualRentForVacantNumbers(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscnumbers2LetNumbers")) {
                                        Double number = biis2gif_Double(cell);
                                        if (number != null)
                                            val.setRentalSituationMiscnumbers2LetNumbers(number.intValue());

                                    } else if (biisKeyName.equals("RentalSituationMiscnumbers2ContractualAnnualRent")) {
                                        val.setRentalSituationMiscnumbers2ContractualAnnualRent(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscnumbers2EstimatedAnnualRentForLetNumbers")) {
                                        val.setRentalSituationMiscnumbers2EstimatedAnnualRentForLetNumbers(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("RentalSituationMiscnumbers2VacantNumbers")) {
                                        Double number = biis2gif_Double(cell);
                                        if (number != null)
                                            val.setRentalSituationMiscnumbers2VacantNumbers(number.intValue());

                                    } else if (biisKeyName.equals("RentalSituationMiscnumbers2EstimatedAnnualRentForVacantNumbers")) {
                                        val.setRentalSituationMiscnumbers2EstimatedAnnualRentForVacantNumbers(biis2gif_Amount(cell, mainCurrency));

                                    } else if (biisKeyName.equals("DataSupplierNumber")) {
                                        val.setExpertId(getCellStringValue(cell));
                                        setObjectId(prop, val);

                                    }
                                } catch (Throwable t) {
                                    logger.warn("cannot convert '" + biisKeyName + "' of cell [" + CellReference.convertNumToColString(columnIndex)
                                            + (rowIndex + 1) + "], value='" + getCellValueObject(cell) + "'", t);

                                    t.printStackTrace();

                                    if (biisKeyName.equals("Currency") && mainCurrency == null) {
                                        throw new Exception("error - cannot find mainCurrency", t);
                                    }
                                    if (biisKeyName.equals("ArealUnit") && mainAreaMeasurement == null) {
                                        throw new Exception("error - cannot find mainAreaMeasurement", t);
                                    }
                                }
                            }
                        }
                    }

                    if (rowIndex != 0) {
                        if (prop.getLabel() == null) {
                            StringBuilder propLabel = new StringBuilder();
                            propLabel.append(valAddress.getStreet());
                            if (valAddress.getHousenumber() != null) {
                                propLabel.append(" ");
                                propLabel.append(valAddress.getHousenumber());
                            }
                            propLabel.append(", ");
                            propLabel.append(valAddress.getZip());
                            propLabel.append(" ");
                            propLabel.append(valAddress.getCity());
                            prop.setLabel(propLabel.toString());
                        }

                        if (prop.getValuations() == null || prop.getValuations().size() == 0) {
                            logger.error("cannot append valuation for property in row " + (rowIndex + 1) + ". IDs correct?");
                            throw new Exception("cannot append valuation for property. IDs correct?");
                        }
                        
                        new ValuationValidator().validate(val);
                    }
                } catch (Throwable t) {
                    logger.error("cannot read row " + (rowIndex + 1), t);
                }
            }
        } catch (Exception e) {
            logger.error(e);
        } finally {
            try {
                workbook.close();
            } catch (Throwable e) {
            }
        }
    }

    @Override
    public ImportWorkerConfiguration getRequiredConfigurationArguments() {
        return new ImportWorkerConfiguration() {
            {
                SortedMap<String, InputStream> streams = getStreams();
                streams.put(PARAMETER_NAME_STREAM, null);

                SortedMap<String, String> strings = getStrings();
                strings.put(PARAMETER_NAME_SHEET_NAME, null);

                SortedMap<String, Integer> integers = getIntegers();
                integers.put(PARAMETER_NAME_SHEET_IDX, null);
            }
        };
    }

    @Override
    public PluginComponent<ImportWorkerConfiguration> getConfigGui() {
        // null => DefaultConfigGui
        return null;
    }

    @Override
    public Container getContainer() {
        return container;
    }

    private void setObjectId(Property prop, Valuation val) {
        if (val.getExpertId() != null && val.getValidFrom() != null) {
            String valId = val.getExpertId() + "_" + new SimpleDateFormat("yyyy-MM-dd").format(val.getValidFrom().toDate());
            val.setObjectIdSender(valId);

            Map<String, Valuation> valuations = prop.getValuations();
            if (valuations == null) {
                valuations = new HashMap<String, Valuation>();
            }
            valuations.put(valId, val);
            prop.setValuations(valuations);
        }
    }

    private AreaMeasurement biis2gif_AreaMeasureMent(String biisValue) {
        if (biisValue == null)
            return null;

        if (biisValue.equals("sqft")) {
            return AreaMeasurement.SQFT;
        } else if (biisValue.equals("qm")) {
            return AreaMeasurement.SQM;
        } else if (biisValue.equals("tsubo") || biisValue.equals("pyeong")) {
            return AreaMeasurement.TSUBO;
        } else {
            return AreaMeasurement.NOT_SPECIFIED;
        }
    }

    private LocalDate biis2gif_Date(Cell cell) {
        if (cell == null)
            return null;

        Date dateVal = cell.getDateCellValue();

        if (dateVal == null)
            return null;

        return LocalDate.fromDateFields(dateVal);
    }

    private LocalDate biis2gif_Year(Cell cell) {
        if (cell == null)
            return null;

        String yearStr = getCellStringValue(cell);
        if (yearStr == null)
            return null;

        if (!Pattern.matches("[0-9.-]+", yearStr)) {
            throw new Error("cell value doesn't matches pattern: [0-9.-]+");
        }

        return LocalDate.parse(yearStr.substring(0, 4) + "-01-01");
    }

    private Boolean biis2gif_Boolean(Cell cell) {
        if (cell == null)
            return null;

        if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return cell.getBooleanCellValue();
        } else {
            String strVal = getCellStringValue(cell);
            if (strVal.toUpperCase().equals("TRUE")) {
                return true;
            }
            if (strVal.toUpperCase().equals("FALSE")) {
                return false;
            }
        }

        return null;
    }

    private Area biis2gif_Area(Cell cell, AreaMeasurement areaMeasurement) {
        Double val = biis2gif_Double(cell);
        if (val == null)
            return null;

        return new Area(val, areaMeasurement, AreaType.NOT_SPECIFIED);
    }

    private Amount biis2gif_Amount(Cell cell, Currency currency) {
        Double val = biis2gif_Double(cell);
        if (val == null)
            return null;

        return new Amount(val, currency);
    }

    private Double biis2gif_Double(Cell cell) {
        if (cell == null)
            return null;

        if (cell.getCellType() != Cell.CELL_TYPE_NUMERIC)
            return null;

        return cell.getNumericCellValue();
    }

    private Period biis2gif_Period(Cell cell) {
        Double value = biis2gif_Double(cell);
        
        return Period.years(value.intValue());
    }

    private ConstructionPhase biis2gif_ConstructionPhase(String biisValue) {
        if (biisValue == null)
            return null;

        if (biisValue.equals("F")) {
            return ConstructionPhase.COMPLETED;
        } else if (biisValue.equals("I")) {
            return ConstructionPhase.IN_COMPLETION;
        } else if (biisValue.equals("P")) {
            return ConstructionPhase.PLANNED;
        } else if (biisValue.substring(0, 1).equals("0")) {
            return ConstructionPhase.OTHER;
        } else {
            throw new Error("unknown enumeration value");
        }
    }

    private ValuationType1 biis2gif_ValuationType1(String biisValue) {
        if (biisValue == null)
            return null;

        if (biisValue.equals("Fondsgutachten")) {
            return ValuationType1.FUND;
        } else if (biisValue.equals("Privatgutachten")) {
            return ValuationType1.PRIVATE;
        } else if (biisValue.equals("Gerichtsgutachten")) {
            return ValuationType1.COURT;
        } else if (biisValue.equals("Fremdgutachten")) {
            return ValuationType1.THIRD_PERSON;
        } else {
            throw new Error("unknown enumeration value");
        }
    }

    private ValuationType2 biis2gif_ValuationType2(String biisValue) {
        if (biisValue == null)
            return null;

        if (biisValue.equals("U")) {
            return ValuationType2.UNKNOWN;
        } else if (biisValue.equals("E")) {
            return ValuationType2.FIRST_VALUATION;
        } else if (biisValue.equals("N")) {
            return ValuationType2.REVALUATION;
        } else if (biisValue.equals("V")) {
            return ValuationType2.MARKET_VALUATION_REPORT;
        } else {
            throw new Error("unknown enumeration value");
        }
    }

    private UseType biis2gif_UseType(String biisValue) {
        if (biisValue == null)
            return null;

        if (biisValue.equals("Buero")) {
            return UseType.OFFICE;
        } else if (biisValue.equals("Handel")) {
            return UseType.RETAIL;
        } else if (biisValue.equals("Industrie(Lager,Hallen)")) {
            return UseType.INDUSTRY;
        } else if (biisValue.equals("Keller/Archiv")) {
            return UseType.OTHER;
        } else if (biisValue.equals("Gastronomie")) {
            return UseType.GASTRONOMY;
        } else if (biisValue.equals("Hotel")) {
            return UseType.HOTEL;
        } else if (biisValue.equals("Wohnen")) {
            return UseType.RESIDENTIAL;
        } else if (biisValue.equals("Freizeit")) {
            return UseType.LEISURE;
        } else if (biisValue.equals("Garage/TG")) {
            return UseType.PARKING;
        } else if (biisValue.equals("Aussenstellplaetze")) {
            return UseType.PARKING;
        } else if (biisValue.equals("unbekannt")) {
            return UseType.NOT_SPECIFIED;
        } else {
            throw new Error("unknown enumeration value");
        }
    }

    private OwnershipType biis2gif_OwnershipType(Cell cell) {
        if (cell == null)
            return null;

        String biisValue = getCellStringValue(cell).substring(0, 1);
        if (biisValue == null)
            return null;

        if (biisValue.equals("0")) { // unbekannt
            return OwnershipType.OTHER;
        } else if (biisValue.equals("1")) { // Dingliches Nutzungsrecht
            return OwnershipType.OTHER;
        } else if (biisValue.equals("2")) { // Erbbaurecht
            return OwnershipType.LEASEHOLD;
        } else if (biisValue.equals("3")) { // gemischte Eigentumsform
            return OwnershipType.OTHER;
        } else if (biisValue.equals("4")) { // Teileigentum
            return OwnershipType.OTHER;
        } else if (biisValue.equals("5")) { // Volleigentum
            return OwnershipType.FREEHOLDER;
        } else if (biisValue.equals("6")) { // Volumeneigentum
            return OwnershipType.OTHER;
        } else {
            throw new Error("unknown enumeration value");
        }
    }

    private RetailLocationType biis2gif_RetailLocationType(String biisValue) {
        if (biisValue == null)
            return null;

        if (biisValue.equals("1a")) {
            return RetailLocationType.HIGH_STREET;
        } else if (biisValue.equals("1b")) {
            return RetailLocationType.CITY_CENTRE_OTHER;
        } else if (biisValue.equals("2a")) {
            return RetailLocationType.MAJOR_ROUTE;
        } else if (biisValue.equals("2b")) {
            return RetailLocationType.SUBURBAN_OTHER;
        } else if (biisValue.equals("c")) {
            return RetailLocationType.NON_URBAN;
        } else if (biisValue.equals("(unbekannt)") || biisValue.equals("unbekannt")) {
            return RetailLocationType.UNKNOWN;
        } else {
            throw new Error("unknown enumeration value");
        }
    }

    private ObjectCondition biis2gif_Condition(String biisValue) {
        if (biisValue == null)
            return null;

        if (biisValue.equals("sehr gut")) {
            return ObjectCondition.NEW;
        } else if (biisValue.equals("gut")) {
            return ObjectCondition.AGE_APPROPRIATE;
        } else if (biisValue.equals("durchschnittlich")) {
            return ObjectCondition.AGE_APPROPRIATE;
        } else if (biisValue.equals("schlecht")) {
            return ObjectCondition.IN_NEED_OF_REPAIR;
        } else if (biisValue.equals("(unbekannt)")) {
            return ObjectCondition.NOT_AVAILABLE;
        } else {
            throw new Error("unknown enumeration value");
        }
    }

    private InteriorQuality biis2gif_InteriorQuality(String biisValue) {
        if (biisValue == null)
            return null;

        if (biisValue.equals("stark gehoben")) {
            return InteriorQuality.LUXURY;
        } else if (biisValue.equals("gehoben")) {
            return InteriorQuality.SOPHISTICATED;
        } else if (biisValue.equals("mittel")) {
            return InteriorQuality.NORMAL;
        } else if (biisValue.equals("einfach")) {
            return InteriorQuality.SIMPLE;
        } else if (biisValue.equals("(unbekannt)")) {
            return InteriorQuality.SIMPLE;
        } else {
            throw new Error("unknown enumeration value");
        }
    }

    private String getCellStringValue(Cell cell) {
        Object obj = getCellValueObject(cell);
        if (obj == null)
            return null;

        return obj.toString();
    }

    private Object getCellValueObject(Cell cell) {
        switch (cell.getCellType()) {
        case Cell.CELL_TYPE_STRING:
            return cell.getRichStringCellValue();
        case Cell.CELL_TYPE_NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            } else {
                return cell.getNumericCellValue();
            }
        case Cell.CELL_TYPE_BOOLEAN:
            return cell.getBooleanCellValue();
        case Cell.CELL_TYPE_FORMULA:
            return cell.getCellFormula();
        default:
            return null;
        }
    }
}
