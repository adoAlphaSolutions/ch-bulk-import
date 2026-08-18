// ============================================================================
// Toll Product Import Generator — Content Hub External Component (multi-category)
// ----------------------------------------------------------------------------
// Browser replacement for the Excel-macro "Export to Content Hub File" flow.
// A single Category dropdown drives everything: which intake field-map to use,
// any special build logic, the update match strategy, and the category value.
//
// Supported categories:
//   Tile, Carpet, Flooring (Hardwood/LVP/Laminate), Cabinets — Vanities,
//   Cabinets — Hardware, Cabinets — Enhancements (Cabinets are create-only).
//
// Input : the vendor INTAKE spreadsheet (row 1 = friendly headers). Add
//         id/identifier columns for updates (categories that support it).
// Output: a workbook with one sheet named "M.PCM.Product", option-list display
//         values resolved to identifiers. Nothing is written to Content Hub.
//
// Option lists are read live from Content Hub data sources when available and
// cached; an embedded snapshot (incl. the TB.Division and TB.PCM.Category
// taxonomies, which the data source client can't read) is the fallback.
// ============================================================================

const SHEET_NAME = 'M.PCM.Product';

// ---- Embedded option-list snapshot (fallback): source -> [[identifier, label], ...]
const DEFAULT_LOOKUPS = {"TB.PCM.Manufacturer": [["TB.PCM.Manufacturer.Amerock", "Amerock"], ["TB.PCM.Manufacturer.Anderson", "Anderson"], ["TB.PCM.Manufacturer.ArizonaTile", "Arizona Tile"], ["TB.PCM.Manufacturer.BerensonHardware", "Berenson Hardware"], ["TB.PCM.Manufacturer.CaliforniaMantelFireplace", "California Mantel & Fireplace"], ["TB.PCM.Manufacturer.Consentino", "Consentino"], ["TB.PCM.Manufacturer.Cove", "Cove"], ["TB.PCM.Manufacturer.DalTile", "DalTile"], ["TB.PCM.Manufacturer.ElectricMirror", "Electric Mirror"], ["TB.PCM.Manufacturer.Emser", "Emser"], ["TB.PCM.Manufacturer.Emtek", "Emtek"], ["TB.PCM.Manufacturer.HomeSiteServices", "Home Site Services"], ["TB.PCM.Manufacturer.HunterDouglas", "Hunter Douglas"], ["TB.PCM.Manufacturer.Infratech", "Infratech"], ["TB.PCM.Manufacturer.JamisonCollections", "Jamison Collections"], ["TB.PCM.Manufacturer.JennAir", "JennAir"], ["TB.PCM.Manufacturer.KitchenAid", "KitchenAid"], ["TB.PCM.Manufacturer.Kohler", "Kohler"], ["TB.PCM.Manufacturer.Kwikset", "Kwikset"], ["TB.PCM.Manufacturer.Marazzi", "Marazzi"], ["TB.PCM.Manufacturer.Metropolitan", "Metropolitan"], ["TB.PCM.Manufacturer.Mohawk", "Mohawk"], ["TB.PCM.Manufacturer.PrecisionCabinetry", "Precision Cabinetry"], ["TB.PCM.Manufacturer.ProgressLighting", "Progress Lighting"], ["TB.PCM.Manufacturer.Provenza", "Provenza"], ["TB.PCM.Manufacturer.Shaw", "Shaw"], ["TB.PCM.Manufacturer.SherwinWilliams", "Sherwin Williams"], ["TB.PCM.Manufacturer.Sterling", "Sterling"], ["TB.PCM.Manufacturer.SubZero", "Sub-Zero"], ["TB.PCM.Manufacturer.TopKnobs", "Top Knobs"], ["TB.PCM.Manufacturer.Tuftex", "Tuftex"], ["TB.PCM.Manufacturer.Unbranded", "Unbranded"], ["TB.PCM.Manufacturer.WesternWindowSystems", "Western Window Systems"], ["TB.PCM.Manufacturer.Whirlpool", "Whirlpool"], ["TB.PCM.Manufacturer.Wolf", "Wolf"], ["TB.PCM.Manufacturer.CenturyCabinetry", "Century Cabinetry"], ["TB.PCM.Manufacturer.AmericanOlean", "American Olean"], ["TB.PCM.Manufacturer.UrbanEffects", "Urban Effects"], ["TB.PCM.Manufacturer.Coyote", "Coyote"], ["TB.PCM.Manufacturer.MSI", "MSI"], ["TB.PCM.Manufacturer.ASKO", "ASKO"], ["TB.PCM.Manufacturer.Mullican", "Mullican"], ["TB.PCM.Manufacturer.Mannington", "Mannington"]], "TB.PCM.Brand": [["TB.PCM.Brand.AmericanOlean", "American Olean"], ["TB.PCM.Brand.ArizonaTile", "Arizona Tile"], ["TB.PCM.Brand.Daltile", "Daltile"], ["TB.PCM.Brand.EmserTile", "Emser Tile"], ["TB.PCM.Brand.Marazzi", "Marazzi"], ["TB.PCM.Brand.Shaw", "Shaw"]], "TB.PCM.UnitOfMeasure": [["TB.PCM.UnitOfMeasure.SQF", "SQF"], ["TB.PCM.UnitOfMeasure.Each", "Each"]], "TB.PCM.TileColorFamily": [["TB.PCM.TileColorFamily.Beige", "Beige"], ["TB.PCM.TileColorFamily.Black", "Black"], ["TB.PCM.TileColorFamily.Blue", "Blue"], ["TB.PCM.TileColorFamily.Brown", "Brown"], ["TB.PCM.TileColorFamily.Cream", "Cream"], ["TB.PCM.TileColorFamily.Gray", "Gray"], ["TB.PCM.TileColorFamily.Green", "Green"], ["TB.PCM.TileColorFamily.Multicolor", "Multicolor"], ["TB.PCM.TileColorFamily.Orange", "Orange"], ["TB.PCM.TileColorFamily.Pink", "Pink"], ["TB.PCM.TileColorFamily.Red", "Red"], ["TB.PCM.TileColorFamily.Taupe", "Taupe"], ["TB.PCM.TileColorFamily.TerraCotta", "Terra Cotta"], ["TB.PCM.TileColorFamily.White", "White"], ["TB.PCM.TileColorFamily.Yellow", "Yellow"]], "TB.PCM.TileStyle": [["TB.PCM.TileStyle.Brick", "Brick"], ["TB.PCM.TileStyle.ConcreteLook", "Concrete Look"], ["TB.PCM.TileStyle.Fabric", "Fabric"], ["TB.PCM.TileStyle.Handmade", "Handmade"], ["TB.PCM.TileStyle.MarbleLook", "Marble Look"], ["TB.PCM.TileStyle.Mosaics", "Mosaics"], ["TB.PCM.TileStyle.NaturalStone", "Natural Stone"], ["TB.PCM.TileStyle.Patterned", "Patterned"], ["TB.PCM.TileStyle.Solid", "Solid"], ["TB.PCM.TileStyle.StoneLook", "Stone Look"], ["TB.PCM.TileStyle.Textured", "Textured"], ["TB.PCM.TileStyle.WoodLook", "Wood Look"]], "TB.PCM.TileSize": [["TB.PCM.TileSize.HerringboneChevron", "Herringbone/Chevron"], ["TB.PCM.TileSize.Hexagon", "Hexagon"], ["TB.PCM.TileSize.LargeFormat", "Large Format"], ["TB.PCM.TileSize.Pebble", "Pebble"], ["TB.PCM.TileSize.PennyOval", "Penny/Oval"], ["TB.PCM.TileSize.Plank", "Plank"], ["TB.PCM.TileSize.Rectangle", "Rectangle"], ["TB.PCM.TileSize.Square", "Square"], ["TB.PCM.TileSize.Subway", "Subway"], ["TB.PCM.TileSize.UniqueShapes", "Unique Shapes"]], "TB.PCM.TileMaterial": [["TB.PCM.TileMaterial.Ceramic", "Ceramic"], ["TB.PCM.TileMaterial.Porcelain", "Porcelain"], ["TB.PCM.TileMaterial.Glass", "Glass"], ["TB.PCM.TileMaterial.NaturalStoneMarble", "Natural Stone - Marble"], ["TB.PCM.TileMaterial.NaturalStoneTravertine", "Natural Stone - Travertine"], ["TB.PCM.TileMaterial.NaturalStoneLimestone", "Natural Stone - Limestone"], ["TB.PCM.TileMaterial.NaturalStoneMixed", "Natural Stone - Mixed"]], "TB.PCM.TileFinish": [["TB.PCM.TileFinish.3DSculptural", "3D/Sculptural"], ["TB.PCM.TileFinish.Glossy", "Glossy"], ["TB.PCM.TileFinish.Matte", "Matte"], ["TB.PCM.TileFinish.Satin", "Satin"], ["TB.PCM.TileFinish.SlipResistant", "Slip Resistant"], ["TB.PCM.TileFinish.Smooth", "Smooth"], ["TB.PCM.TileFinish.Textured", "Textured"]], "TB.PCM.CarpetFlooringColorFamily": [["TB.PCM.CarpetFlooringColorFamily.White", "White"], ["TB.PCM.CarpetFlooringColorFamily.Cream", "Cream"], ["TB.PCM.CarpetFlooringColorFamily.Beige", "Beige"], ["TB.PCM.CarpetFlooringColorFamily.RedPink", "Red / Pink"], ["TB.PCM.CarpetFlooringColorFamily.Orange", "Orange"], ["TB.PCM.CarpetFlooringColorFamily.Yellow", "Yellow"], ["TB.PCM.CarpetFlooringColorFamily.Taupe", "Taupe"], ["TB.PCM.CarpetFlooringColorFamily.Brown", "Brown"], ["TB.PCM.CarpetFlooringColorFamily.Gray", "Gray"], ["TB.PCM.CarpetFlooringColorFamily.Charcoal", "Charcoal"], ["TB.PCM.CarpetFlooringColorFamily.Multicolor", "Multicolor"], ["TB.PCM.CarpetFlooringColorFamily.Green", "Green"], ["TB.PCM.CarpetFlooringColorFamily.Blue", "Blue"]], "TB.PCM.CarpetFlooringFamilyTones": [["TB.PCM.CarpetFlooringFamilyTones.Blue", "Blue"], ["TB.PCM.CarpetFlooringFamilyTones.Brown", "Brown"], ["TB.PCM.CarpetFlooringFamilyTones.Cream", "Cream"], ["TB.PCM.CarpetFlooringFamilyTones.Gray", "Gray"], ["TB.PCM.CarpetFlooringFamilyTones.Taupe", "Taupe"], ["TB.PCM.CarpetFlooringFamilyTones.WarmNeutral", "Warm Neutral"], ["TB.PCM.CarpetFlooringFamilyTones.White", "White"]], "TB.PCM.CarpetFlooringMaterial": [["TB.PCM.CarpetFlooringMaterial.Nylon", "Nylon"], ["TB.PCM.CarpetFlooringMaterial.Polyester", "Polyester"], ["TB.PCM.CarpetFlooringMaterial.Wool", "Wool"], ["TB.PCM.CarpetFlooringMaterial.EndurancePET", "Endurance PET"], ["TB.PCM.CarpetFlooringMaterial.EndurancewColorguard", "Endurance w Colorguard"], ["TB.PCM.CarpetFlooringMaterial.SolutionDyedPET", "Solution Dyed PET"], ["TB.PCM.CarpetFlooringMaterial.Anso", "ANSO High Performance Nylon"], ["TB.PCM.CarpetFlooringMaterial.ANSOEcoNylon", "ANSO Eco Nylon"]], "TB.PCM.CarpetFlooringStyle": [["TB.PCM.CarpetFlooringStyle.CutPile", "Cut Pile"], ["TB.PCM.CarpetFlooringStyle.LoopPile", "Loop Pile"], ["TB.PCM.CarpetFlooringStyle.Patterned", "Patterned"]], "TB.PCM.FlooringColorFamily": [["TB.PCM.FlooringColorFamily.Blonde", "Blonde"], ["TB.PCM.FlooringColorFamily.Honey", "Honey"], ["TB.PCM.FlooringColorFamily.Greige", "Greige"], ["TB.PCM.FlooringColorFamily.Gray", "Gray"], ["TB.PCM.FlooringColorFamily.DarkUmber", "Dark/Umber"], ["TB.PCM.FlooringColorFamily.White", "White"]], "TB.PCM.ColorVariation": [["TB.PCM.ColorVariation.Low", "Low"], ["TB.PCM.ColorVariation.Moderate", "Moderate"], ["TB.PCM.ColorVariation.High", "High"]], "TB.PCM.FlooringSpecies": [["TB.PCM.FlooringSpecies.WhiteOak", "White Oak"], ["TB.PCM.FlooringSpecies.RedOak", "Red Oak"], ["TB.PCM.FlooringSpecies.Hickory", "Hickory"], ["TB.PCM.FlooringSpecies.Ash", "Ash"], ["TB.PCM.FlooringSpecies.Pecan", "Pecan"], ["TB.PCM.FlooringSpecies.Maple", "Maple"], ["TB.PCM.FlooringSpecies.Walnut", "Walnut"]], "TB.PCM.SurfaceTextureVisual": [["TB.PCM.SurfaceTextureVisual.Smooth", "Smooth"], ["TB.PCM.SurfaceTextureVisual.ReclaimedDistressed", "Reclaimed/Distressed"], ["TB.PCM.SurfaceTextureVisual.WireBrushed", "Wire Brushed"], ["TB.PCM.SurfaceTextureVisual.HandScraped", "Hand Scraped"], ["TB.PCM.SurfaceTextureVisual.SawnFaced", "Sawn Faced"]], "TB.PCM.CabinetWoodType": [["TB.PCM.CabinetWoodType.Beech", "Beech"], ["TB.PCM.CabinetWoodType.Cherry", "Cherry"], ["TB.PCM.CabinetWoodType.Hickory", "Hickory"], ["TB.PCM.CabinetWoodType.Maple", "Maple"], ["TB.PCM.CabinetWoodType.Oak", "Oak"], ["TB.PCM.CabinetWoodType.Walnut", "Walnut"]], "TB.PCM.CabinetColorFamily": [["TB.PCM.CabinetColorFamily.Black", "Black"], ["TB.PCM.CabinetColorFamily.Blonde", "Blonde"], ["TB.PCM.CabinetColorFamily.Blue", "Blue"], ["TB.PCM.CabinetColorFamily.Brown", "Brown"], ["TB.PCM.CabinetColorFamily.Dark", "Dark"], ["TB.PCM.CabinetColorFamily.DarkWood", "Dark Wood"], ["TB.PCM.CabinetColorFamily.Gray", "Gray"], ["TB.PCM.CabinetColorFamily.GrayWood", "Gray Wood"], ["TB.PCM.CabinetColorFamily.Green", "Green"], ["TB.PCM.CabinetColorFamily.Greige", "Greige"], ["TB.PCM.CabinetColorFamily.Honey", "Honey"], ["TB.PCM.CabinetColorFamily.LightWood", "Light Wood"], ["TB.PCM.CabinetColorFamily.MediumWood", "Medium Wood"], ["TB.PCM.CabinetColorFamily.White", "White"]], "TB.PCM.CabinetFinish": [["TB.PCM.CabinetFinish.HighGloss", "High Gloss"], ["TB.PCM.CabinetFinish.PaintGrade", "Painted"], ["TB.PCM.CabinetFinish.Stained", "Stained"], ["TB.PCM.CabinetFinish.Textured", "Textured"]], "TB.PCM.CabinetHardwareType": [["TB.PCM.CabinetHardwareType.knob", "knob"], ["TB.PCM.CabinetHardwareType.pull", "pull"]], "TB.PCM.CabinetEnhancementsType": [["TB.PCM.CabinetEnhType.DTB3Y", "Kitchen Enhancement Options"], ["TB.PCM.CabinetEnhType.BOpt", "Bath Enhancement Options"], ["TB.PCM.CabinetEnhType.HOpt", "Whole House Enhancement Options"]], "TB.Division": [["TB.Division.AZR", "Arizona"], ["TB.Division.CA-LA", "CA-LOS ANGELES"], ["TB.Division.CA-N", "CA-NORTHERN"], ["TB.Division.CA-S", "CA-SOUTHERN"], ["TB.Division.CA-SAC", "CA-SACRAMENTO"], ["TB.Division.CAU", "Caifornia City Living"], ["TB.Division.CO-S", "COLORADO SPRINGS"], ["TB.Division.COL", "Colorado"], ["TB.Division.DCU", "DC City Living"], ["TB.Division.FL-C", "FL-CENTRAL"], ["TB.Division.FL-E", "FL-EAST"], ["TB.Division.FL-N", "FL-NORTH"], ["TB.Division.FL-TS", "FL-TAMPA SARASOTA"], ["TB.Division.FL-W", "FL-WEST"], ["TB.Division.GA-A", "GEORGIA-ATLANTA"], ["TB.Division.GAT", "Georgia Metro"], ["TB.Division.IL", "ILLINOIS"], ["TB.Division.ILL", "Illinois"], ["TB.Division.IT", "IT Test Division"], ["TB.Division.MA", "MASSACHUSETTS"], ["TB.Division.MAS", "Massachusetts"], ["TB.Division.MD-MS", "MARYLAND-MARYLAND SHORE"], ["TB.Division.MI", "MICHIGAN"], ["TB.Division.MIC", "Michigan"], ["TB.Division.MN", "MINNESOTA"], ["TB.Division.MNN", "Minnesota"], ["TB.Division.NCC", "North Carolina Charlotte"], ["TB.Division.NCR", "North Carolina Raleigh"], ["TB.Division.NJ", "NEW JERSEY"], ["TB.Division.NJS", "New Jersey Suburban"], ["TB.Division.NotDefined", "Not Defined"], ["TB.Division.NV-LV", "NV-LAS VEGAS"], ["TB.Division.NV-R", "NV-RENO"], ["TB.Division.NY", "NEW YORK"], ["TB.Division.NYL", "New York Long Island"], ["TB.Division.NYS", "New York Suburban"], ["TB.Division.NYU", "NY City Living"], ["TB.Division.OR", "OREGON"], ["TB.Division.PA", "PENNSYLVANIA"], ["TB.Division.PAN", "Pennsylvania North"], ["TB.Division.PHU", "Philadelphia City Living"], ["TB.Division.PSN", "Philadelphia Suburbs North"], ["TB.Division.TN", "TENNESSEE"], ["TB.Division.TX-AUS", "TX-AUSTIN"], ["TB.Division.TX-DFW", "TX-DALLAS"], ["TB.Division.TX-HOU", "TX-HOUSTON"], ["TB.Division.TX-SAN", "TX-SAN ANTONIO"], ["TB.Division.UTA", "Utah"], ["TB.Division.WA-SEA", "WASHINGTON-SEATTLE"], ["TB.Division.WAU", "Washington Seattle City Living"]], "TB.PCM.Category": [["TB.PCM.Category.Appliances", "Appliances"], ["TB.PCM.Category.AudioVideo", "Audio & Video"], ["TB.PCM.Category.Baseboards", "Baseboards"], ["TB.PCM.Category.BaseboardTrim", "Baseboard & Trim"], ["TB.PCM.Category.BathFaucets", "Bath Faucets"], ["TB.PCM.Category.BathSink", "Bath Sink"], ["TB.PCM.Category.BathTubs", "Bath Tubs"], ["TB.PCM.Category.CabinetEnhancements", "Cabinet Enhancements"], ["TB.PCM.Category.CabinetHardware", "Cabinet Hardware"], ["TB.PCM.Category.Cabinets.Cabinets", "Cabinets"], ["TB.PCM.Category.Cabinets", "Cabinets"], ["TB.PCM.Category.CeilingBeam", "Ceiling Beam"], ["TB.PCM.Category.Closets", "Closets"], ["TB.PCM.Category.ColorSchemes", "Color Schemes"], ["TB.PCM.Category.CookingPackage", "Cooking Package"], ["TB.PCM.Category.Countertops", "Countertops"], ["TB.PCM.Category.DecksBalconies", "Decks & Balconies"], ["TB.PCM.Category.DoorHardware", "Door Hardware"], ["TB.PCM.Category.Doors", "Doors"], ["TB.PCM.Category.Drywall", "Drywall"], ["TB.PCM.Category.Electrical", "Electrical"], ["TB.PCM.Category.EnergySolar", "Energy & Solar"], ["TB.PCM.Category.EntertainmentFaucets", "Entertainment Faucets"], ["TB.PCM.Category.EntertainmentSink", "Entertainment Sink"], ["TB.PCM.Category.Exteriors", "Exteriors"], ["TB.PCM.Category.Fencing", "Fencing"], ["TB.PCM.Category.Fireplaces", "Fireplaces"], ["TB.PCM.Category.FireSuppression", "Fire Suppression"], ["TB.PCM.Category.Flooring.Carpet", "Carpet"], ["TB.PCM.Category.Flooring", "Flooring"], ["TB.PCM.Category.Flooring.Laminate", "Laminate"], ["TB.PCM.Category.Flooring.LVP", "LVP"], ["TB.PCM.Category.FloorPlanOptions", "Floor Plan Options"], ["TB.PCM.Category.GasSystems", "Gas Systems"], ["TB.PCM.Category.Gutters", "Gutters"], ["TB.PCM.Category.Hardscaping", "Hardscaping"], ["TB.PCM.Category.Hardwood", "Hardwood"], ["TB.PCM.Category.HomeDet.Fireplace", "Fireplace"], ["TB.PCM.Category.HomeDetails", "Home Details"], ["TB.PCM.Category.HomeTechnology", "Home Technology"], ["TB.PCM.Category.HumidifierSystems", "Humidifier Systems"], ["TB.PCM.Category.HVAC", "HVAC"], ["TB.PCM.Category.Insulation", "Insulation"], ["TB.PCM.Category.Interior", "Interior"], ["TB.PCM.Category.InteriorD.GarageEpoxy", "Garage Epoxy"], ["TB.PCM.Category.InteriorD.Lighting", "Lighting"], ["TB.PCM.Category.InteriorD.Mirrors", "Mirrors"], ["TB.PCM.Category.InteriorDetails", "Interior Details"], ["TB.PCM.Category.KitchenFaucets", "Kitchen Faucets"], ["TB.PCM.Category.KitchenSink", "Kitchen Sink"], ["TB.PCM.Category.Landscaping", "Landscaping"], ["TB.PCM.Category.LaundryAppliances", "Laundry Appliances"], ["TB.PCM.Category.LaundryFaucets", "Laundry Faucets"], ["TB.PCM.Category.LaundrySink", "Laundry Sink"], ["TB.PCM.Category.Network", "Network"], ["TB.PCM.Category.Other", "Other"], ["TB.PCM.Category.Outdoor", "Outdoor"], ["TB.PCM.Category.OutdoorAccessories", "Outdoor Accessories"], ["TB.PCM.Category.OutdoorFireplace", "Outdoor Fireplace"], ["TB.PCM.Category.OutdoorKitchen", "Outdoor Kitchen"], ["TB.PCM.Category.OutdoorOptions", "Outdoor Options"], ["TB.PCM.Category.Paint", "Paint"], ["TB.PCM.Category.PerformanceShowering", "Performance Showering"], ["TB.PCM.Category.Plumbing.BathHardware", "Bath Hardware"], ["TB.PCM.Category.Plumbing", "Plumbing"], ["TB.PCM.Category.Plumbing.ShowerEnclosures", "Shower Enclosures"], ["TB.PCM.Category.Plumbing.Toilets", "Toilets"], ["TB.PCM.Category.Plumbing.Tubs", "Tubs"], ["TB.PCM.Category.Pool", "Pool"], ["TB.PCM.Category.Refrigerators", "Refrigerators"], ["TB.PCM.Category.RoughPlumbing", "Rough Plumbing"], ["TB.PCM.Category.Security", "Security"], ["TB.PCM.Category.ShowerPackage", "Shower Package"], ["TB.PCM.Category.Sinks", "Sinks"], ["TB.PCM.Category.SmartLighting", "Smart Lighting"], ["TB.PCM.Category.SoapDispenser", "Soap Dispenser"], ["TB.PCM.Category.SolidSurfaces", "Solid Surfaces"], ["TB.PCM.Category.Stairs", "Stairs"], ["TB.PCM.Category.Tiles", "Tile"], ["TB.PCM.Category.Trim", "Trim"], ["TB.PCM.Category.TubFaucets", "Tub Faucets"], ["TB.PCM.Category.Tubs", "Tubs"], ["TB.PCM.Category.TubShowerPackage", "Tub/Shower Package"], ["TB.PCM.Category.UnderCabinetRefrigeration", "Under Cabinet Refrigeration"], ["TB.PCM.Category.Wallfinish", "Wall finish"], ["TB.PCM.Category.WaterSystems", "Water Systems"], ["TB.PCM.Category.WindowCoverings", "Window Coverings"]], "TB.PCM.CarpetFlooringColorTone": [["TB.PCM.CarpetFlooringColorTone.Light", "Light"], ["TB.PCM.CarpetFlooringColorTone.Medium", "Medium"], ["TB.PCM.CarpetFlooringColorTone.Dark", "Dark"]], "TB.PCM.CarpetFlooringSpecialFeatures": [["TB.PCM.CarpetFlooringSpecialFeatures.StainRes", "Inherently Stain Resistant"], ["TB.PCM.CarpetFlooringSpecialFeatures.R2X", "R2X Stain Protection"], ["TB.PCM.CarpetFlooringSpecialFeatures.Colorguard", "Colorguard"]], "TB.PCM.CarpetFlooringCarpetStyle": [["TB.PCM.CarpetFlooringCarpetStyle.Pattern", "Pattern"], ["TB.PCM.CarpetFlooringCarpetStyle.TextureSolid", "Texture Solid"], ["TB.PCM.CarpetFlooringCarpetStyle.TextureTonal", "Texture Tonal"], ["TB.PCM.CarpetFlooringCarpetStyle.TextureTweed", "Texture Tweed"], ["TB.PCM.CarpetFlooringCarpetStyle.TonalPattern", "Tonal Pattern"], ["TB.PCM.CarpetFlooringCarpetStyle.SolidColorPattern", "Solid Color Pattern"], ["TB.PCM.CarpetFlooringCarpetStyle.LoopPattern", "Loop Pattern"], ["TB.PCM.CarpetFlooringCarpetStyle.LCLPattern", "LCL Pattern"], ["TB.PCM.CarpetFlooringCarpetStyle.Solid", "Solid"]], "TB.PCM.CarpetFlooringConstruction": [["TB.PCM.CarpetFlooringConstruction.CutPile", "Cut Pile"], ["TB.PCM.CarpetFlooringConstruction.CutLoop", "Cut & Loop"], ["TB.PCM.CarpetFlooringConstruction.LoopPile", "Loop Pile"]]};

let liveCache = {};
const DATA_SOURCE_NAMES = [
  'TB.PCM.Manufacturer','TB.PCM.Brand','TB.PCM.UnitOfMeasure','TB.PCM.TileColorFamily','TB.PCM.TileStyle',
  'TB.PCM.TileSize','TB.PCM.TileMaterial','TB.PCM.TileFinish','TB.PCM.CarpetFlooringColorFamily',
  'TB.PCM.CarpetFlooringMaterial','TB.PCM.CarpetFlooringStyle','TB.PCM.FlooringColorFamily','TB.PCM.ColorVariation',
  'TB.PCM.FlooringSpecies','TB.PCM.SurfaceTextureVisual','TB.PCM.CabinetWoodType','TB.PCM.CabinetColorFamily',
  'TB.PCM.CabinetFinish','TB.PCM.CabinetHardwareType','TB.PCM.CabinetEnhancementsType'
];

// Field -> backing option-list source (mirrors GetLookupSourceForCHField).
const FIELD_SOURCE = {
  'TB.PCM.Product.Manufacturer': 'TB.PCM.Manufacturer',
  'TB.PCM.Product.TileUnitOfMeasure': 'TB.PCM.UnitOfMeasure',
  'TB.PCM.DivisionSelected': 'TB.Division'
};
function sourceFor(field) { return FIELD_SOURCE[field] || field; }

const ID_LABELS = { 'id': 'id', 'content hub id': 'id', 'identifier': 'identifier', 'content hub identifier': 'identifier' };

// ---- Flooring shared config (one sheet per sub-category in the workbook) ----
const FLOORING_FIELDMAP = {
  'item #': 'TB.PCM.E1ItemNumber', 'vendor sku (if available)': 'TB.PCM.Product.SKU', 'manufacturer': 'TB.PCM.Product.Manufacturer',
  'brand': 'TB.PCM.Brand', 'toll family/style name': 'TB.PCM.FamilyName', 'toll style number': 'TB.PCM.Flooring.TollStyleNumber',
  'toll color name': 'Color', 'toll color number': 'TB.PCM.Product.ColorNumber',
  'manufacturer family name (if product is private label)': 'TB.PCM.FlooringManufacturerFamilyName',
  'manufacturer style number (if product is private label)': 'TB.PCM.FlooringManufacturerStyleNumber',
  'manufacturer color name (if product is private label)': 'TB.PCM.FlooringManufacturerColorName',
  'manufacturer color number (if product is private label)': 'TB.PCM.FlooringManufacturerColorNumber',
  'color family': 'TB.PCM.FlooringColorFamily', 'color variation': 'TB.PCM.ColorVariation', 'species': 'TB.PCM.FlooringSpecies',
  'surface texture/visual': 'TB.PCM.SurfaceTextureVisual', 'finish/coating': 'TB.PCM.FinishCoating', 'wear layer (mil)': 'TB.PCM.FlooringWearLayer',
  'thickness': 'TB.PCM.FlooringThickness', 'width (in)': 'TB.PCM.FlooringWidth', 'length (in)': 'TB.PCM.FlooringLength',
  'product description': 'TB.PCM.ProductDescription', 'divisions selected': 'TB.PCM.DivisionSelected'
};
const FLOORING_OPTS = ['TB.PCM.Product.Manufacturer', 'TB.PCM.Brand', 'TB.PCM.FlooringColorFamily', 'TB.PCM.ColorVariation', 'TB.PCM.FlooringSpecies', 'TB.PCM.SurfaceTextureVisual', 'TB.PCM.DivisionSelected'];
const FLOORING_OUT = ['id', 'identifier', 'TB.PCM.Category', 'TB.PCM.ProductName', 'TB.PCM.Product.SKU', 'TB.PCM.E1ItemNumber', 'TB.PCM.TollSKU', 'TB.PCM.Product.Manufacturer', 'TB.PCM.Brand', 'TB.PCM.FamilyName', 'TB.PCM.Flooring.TollStyleNumber', 'Color', 'TB.PCM.Product.ColorNumber', 'TB.PCM.FlooringManufacturerFamilyName', 'TB.PCM.FlooringManufacturerStyleNumber', 'TB.PCM.FlooringManufacturerColorName', 'TB.PCM.FlooringManufacturerColorNumber', 'TB.PCM.FlooringColorFamily', 'TB.PCM.ColorVariation', 'TB.PCM.FlooringSpecies', 'TB.PCM.SurfaceTextureVisual', 'TB.PCM.FinishCoating', 'TB.PCM.FlooringWearLayer', 'TB.PCM.FlooringThickness', 'TB.PCM.FlooringWidth', 'TB.PCM.FlooringLength', 'TB.PCM.ProductDescription', 'TB.PCM.DivisionSelected'];
const FLOORING_MATCH = [['TB.PCM.E1ItemNumber', 'Color'], ['TB.PCM.Product.SKU'], ['TB.PCM.FlooringManufacturerStyleNumber'], ['TB.PCM.Flooring.TollStyleNumber'], ['TB.PCM.FamilyName', 'Color']];
function flooringCfg(label, sheetName, categoryValue) {
  return {
    label, sheetName, sheetMatch: [String(sheetName).toLowerCase()], categoryValue, supportsUpdate: true, matchStrategies: FLOORING_MATCH,
    fieldMap: FLOORING_FIELDMAP, itemCols: null, specialFeatures: null, optionListFields: FLOORING_OPTS,
    outCols: FLOORING_OUT, requiredFields: ['TB.PCM.Category', 'TB.PCM.Product.Manufacturer', 'Color', 'TB.PCM.Product.SKU'], fallbacks: null
  };
}

// ---- Per-category configuration -------------------------------------------
const CATEGORY_CONFIGS = {
  tile: {
    label: 'Tile', categoryValue: 'TB.PCM.Category.Tiles', supportsUpdate: true,
    matchStrategies: [['TB.PCM.E1ItemNumber', 'Color']],
    fieldMap: {
      'product name': 'TB.PCM.ProductName', 'sku': 'TB.PCM.Product.SKU', 'toll sku': 'TB.PCM.TollSKU',
      'manufacturer': 'TB.PCM.Product.Manufacturer', 'brand': 'TB.PCM.Brand', 'family/style name': 'TB.PCM.FamilyName',
      'color': 'Color', 'color family': 'TB.PCM.TileColorFamily', 'look/style': 'TB.PCM.TileStyle',
      'shape': 'TB.PCM.TileSize', 'size': 'TB.PCM.Size', 'material': 'TB.PCM.TileMaterial', 'finish': 'TB.PCM.TileFinish',
      'unit of measure': 'TB.PCM.Product.TileUnitOfMeasure', 'thickness (in)': 'TB.PCM.Product.TileThickness',
      'product description': 'TB.PCM.ProductDescription'
    },
    itemCols: [
      { label: 'floor item #', aoa: 'TB.PCM.AreaOfApplication.FloorTile' },
      { label: 'wall item #', aoa: 'TB.PCM.AreaOfApplication.WallTile' },
      { label: 'backsplash item #', aoa: 'TB.PCM.AreaOfApplication.TileBacksplash' },
      { label: 'shower floor item #', aoa: 'TB.PCM.AreaOfApplication.ShowerFloor' },
      { label: 'listello item #', aoa: 'TB.PCM.AreaOfApplication.Listello' }
    ],
    specialFeatures: null,
    optionListFields: ['TB.PCM.Product.Manufacturer', 'TB.PCM.Brand', 'TB.PCM.TileColorFamily', 'TB.PCM.TileStyle', 'TB.PCM.TileSize', 'TB.PCM.TileMaterial', 'TB.PCM.TileFinish', 'TB.PCM.Product.TileUnitOfMeasure'],
    outCols: ['id', 'identifier', 'TB.PCM.Category', 'TB.PCM.ProductName', 'TB.PCM.Product.SKU', 'TB.PCM.TollSKU', 'TB.PCM.E1ItemNumber', 'TB.PCM.Product.Manufacturer', 'TB.PCM.Brand', 'TB.PCM.FamilyName', 'Color', 'TB.PCM.TileColorFamily', 'TB.PCM.TileStyle', 'TB.PCM.TileSize', 'TB.PCM.Size', 'TB.PCM.TileMaterial', 'TB.PCM.TileFinish', 'TB.PCM.Product.TileUnitOfMeasure', 'TB.PCM.Product.TileThickness', 'TB.PCM.ProductDescription', 'TB.PCM.AreaOfApplication'],
    requiredFields: ['TB.PCM.ProductName', 'TB.PCM.Category', 'TB.PCM.Product.Manufacturer', 'Color', 'TB.PCM.Product.SKU'],
    fallbacks: { nameFrom: ['description 1', 'description 2'], skuFromItems: ['backsplash item #', 'floor item #', 'shower floor item #', 'wall item #', 'listello item #'] }
  },
  carpet: {
    label: 'Carpet', categoryValue: 'TB.PCM.Category.Flooring.Carpet', supportsUpdate: true,
    matchStrategies: [['TB.PCM.E1ItemNumber', 'Color']],
    fieldMap: {
      'item #': 'TB.PCM.E1ItemNumber', 'manufacturer': 'TB.PCM.Product.Manufacturer', 'brand': 'TB.PCM.Brand',
      'toll family/style name': 'TB.PCM.FamilyName', 'toll style number': 'TB.PCM.Flooring.TollStyleNumber',
      'toll color name': 'Color', 'toll color number': 'TB.PCM.Product.ColorNumber',
      'manufacturer name (if product is private label)': 'TB.PCM.FlooringManufacturerFamilyName',
      'manufacturer style number (if product is private label)': 'TB.PCM.FlooringManufacturerStyleNumber',
      'manufacturer color name (if product is private label)': 'TB.PCM.FlooringManufacturerColorName',
      'manufacturer color number (if product is private label)': 'TB.PCM.FlooringManufacturerColorNumber',
      'color family': 'TB.PCM.CarpetFlooringColorFamily', 'material': 'TB.PCM.CarpetFlooringMaterial',
      'style': 'TB.PCM.CarpetFlooringStyle', 'pattern size, if applicable': 'TB.PCM.CarpetFlooringCarpetPatternSize',
      'pattern size': 'TB.PCM.CarpetFlooringCarpetPatternSize',
      'construction': 'TB.PCM.CarpetFlooringConstruction', 'product description': 'TB.PCM.ProductDescription',
      'thickness': 'TB.PCM.FlooringThickness'
    },
    itemCols: null,
    specialFeatures: {
      target: 'TB.PCM.CarpetFlooringSpecialFeatures',
      flags: [
        { label: 'special features- inherently stain resistant', ids: ['TB.PCM.CarpetFlooringSpecialFeatures.StainRes'] },
        { label: 'special features- r2x stain protection', ids: ['TB.PCM.CarpetFlooringSpecialFeatures.R2X'] },
        { label: 'special features- colorguard', ids: ['TB.PCM.CarpetFlooringSpecialFeatures.Colorguard'] },
        { label: 'special features- r2x + colorguard', ids: ['TB.PCM.CarpetFlooringSpecialFeatures.R2X', 'TB.PCM.CarpetFlooringSpecialFeatures.Colorguard'] }
      ]
    },
    optionListFields: ['TB.PCM.Product.Manufacturer', 'TB.PCM.Brand', 'TB.PCM.CarpetFlooringColorFamily', 'TB.PCM.CarpetFlooringMaterial', 'TB.PCM.CarpetFlooringStyle', 'TB.PCM.CarpetFlooringConstruction'],
    outCols: ['id', 'identifier', 'TB.PCM.Category', 'TB.PCM.E1ItemNumber', 'TB.PCM.Product.Manufacturer', 'TB.PCM.Brand', 'TB.PCM.FamilyName', 'TB.PCM.Flooring.TollStyleNumber', 'Color', 'TB.PCM.Product.ColorNumber', 'TB.PCM.FlooringManufacturerFamilyName', 'TB.PCM.FlooringManufacturerStyleNumber', 'TB.PCM.FlooringManufacturerColorName', 'TB.PCM.FlooringManufacturerColorNumber', 'TB.PCM.CarpetFlooringColorFamily', 'TB.PCM.CarpetFlooringMaterial', 'TB.PCM.CarpetFlooringStyle', 'TB.PCM.CarpetFlooringCarpetPatternSize', 'TB.PCM.CarpetFlooringConstruction', 'TB.PCM.FlooringThickness', 'TB.PCM.ProductDescription', 'TB.PCM.CarpetFlooringSpecialFeatures'],
    requiredFields: ['TB.PCM.Category', 'TB.PCM.Product.Manufacturer', 'Color', 'TB.PCM.E1ItemNumber'],
    fallbacks: null
  },
  'flooring-hardwood': flooringCfg('Flooring — Hardwood', 'Hardwood', 'TB.PCM.Category.Hardwood'),
  'flooring-lvp': flooringCfg('Flooring — LVP', 'LVP', 'TB.PCM.Category.Flooring.LVP'),
  'flooring-laminate': flooringCfg('Flooring — Laminate', 'Laminate', 'TB.PCM.Category.Flooring.Laminate'),
  'cabinets-vanities': {
    label: 'Cabinets — Vanities', sheetName: 'Cabinets & Vanities', sheetMatch: ['vanit'], categoryValue: 'TB.PCM.Category.Cabinets.Cabinets', supportsUpdate: false, matchStrategies: [],
    fieldMap: {
      'manufacturer': 'TB.PCM.Product.Manufacturer', 'brand': 'TB.PCM.Brand', 'style/collection name': 'TB.PCM.FamilyName',
      'door material': 'TB.PCM.CabinetWoodType', 'door style': 'TB.PCM.Style', 'color': 'Color', 'color family': 'TB.PCM.CabinetColorFamily',
      'finish': 'TB.PCM.CabinetFinish', 'framed or frameless': 'TB.PCM.CabinetFramedFrameless', 'door construction type': 'TB.PCM.CabinetDoorConstructionType',
      'drawer style': 'TB.PCM.CabinetDrawerStyle', 'product description': 'TB.PCM.ProductDescription', 'divisions selected': 'TB.PCM.DivisionSelected',
      'product name': 'TB.PCM.ProductName', 'mfn sku': 'TB.PCM.Product.SKU', 'toll sku': 'TB.PCM.TollSKU'
    },
    itemCols: null, specialFeatures: null,
    optionListFields: ['TB.PCM.Product.Manufacturer', 'TB.PCM.Brand', 'TB.PCM.CabinetWoodType', 'TB.PCM.CabinetColorFamily', 'TB.PCM.CabinetFinish', 'TB.PCM.DivisionSelected'],
    outCols: ['TB.PCM.Category', 'TB.PCM.ProductName', 'TB.PCM.Product.SKU', 'TB.PCM.TollSKU', 'TB.PCM.Product.Manufacturer', 'TB.PCM.Brand', 'TB.PCM.FamilyName', 'Color', 'TB.PCM.ProductDescription', 'TB.PCM.CabinetWoodType', 'TB.PCM.Style', 'TB.PCM.CabinetColorFamily', 'TB.PCM.CabinetFinish', 'TB.PCM.CabinetFramedFrameless', 'TB.PCM.CabinetDoorConstructionType', 'TB.PCM.CabinetDrawerStyle', 'TB.PCM.DivisionSelected'],
    requiredFields: ['TB.PCM.Category', 'TB.PCM.ProductName', 'TB.PCM.Product.Manufacturer'],
    fallbacks: null
  },
  'cabinets-hardware': {
    label: 'Cabinets — Hardware', sheetName: 'Cabinet Hardware', sheetMatch: ['hardware'], categoryValue: 'TB.PCM.Category.CabinetHardware', supportsUpdate: false, matchStrategies: [],
    fieldMap: {
      'vendor sku': 'TB.PCM.Product.SKU', 'manufacturer': 'TB.PCM.Product.Manufacturer', 'brand': 'TB.PCM.Brand', 'style name': 'TB.PCM.FamilyName',
      'hardware style': 'TB.PCM.CabinetHardwareType', 'finish': 'Color', 'hardware color family': 'TB.PCM.HardwareColorFamily', 'size': 'TB.PCM.Size',
      'product description': 'TB.PCM.ProductDescription', 'product name': 'TB.PCM.ProductName', 'toll sku': 'TB.PCM.TollSKU', 'divisions selected': 'TB.PCM.DivisionSelected'
    },
    itemCols: null, specialFeatures: null,
    optionListFields: ['TB.PCM.Product.Manufacturer', 'TB.PCM.Brand', 'TB.PCM.CabinetHardwareType', 'TB.PCM.DivisionSelected'],
    outCols: ['TB.PCM.Category', 'TB.PCM.ProductName', 'TB.PCM.Product.SKU', 'TB.PCM.TollSKU', 'TB.PCM.Product.Manufacturer', 'TB.PCM.Brand', 'TB.PCM.FamilyName', 'TB.PCM.CabinetHardwareType', 'Color', 'TB.PCM.HardwareColorFamily', 'TB.PCM.Size', 'TB.PCM.ProductDescription', 'TB.PCM.DivisionSelected'],
    requiredFields: ['TB.PCM.Category', 'TB.PCM.ProductName', 'TB.PCM.Product.Manufacturer'],
    fallbacks: null
  },
  'cabinets-enhancements': {
    label: 'Cabinets — Enhancements', sheetName: 'Cabinet Enhancements', sheetMatch: ['enhance'], categoryValue: 'TB.PCM.Category.CabinetEnhancements', supportsUpdate: false, matchStrategies: [],
    fieldMap: {
      'vendor sku': 'TB.PCM.Product.SKU', 'manufacturer': 'TB.PCM.Product.Manufacturer', 'enhancement name': 'TB.PCM.ProductName',
      'enhancement category': 'TB.PCM.CabinetEnhancementsType', 'enhancement location': 'TB.PCM.CabinetEnchancementLocation', 'size': 'TB.PCM.Size',
      'description': 'TB.PCM.ProductDescription', 'toll sku': 'TB.PCM.TollSKU', 'divisions selected': 'TB.PCM.DivisionSelected'
    },
    itemCols: null, specialFeatures: null,
    optionListFields: ['TB.PCM.Product.Manufacturer', 'TB.PCM.CabinetEnhancementsType', 'TB.PCM.DivisionSelected'],
    outCols: ['TB.PCM.Category', 'TB.PCM.ProductName', 'TB.PCM.Product.SKU', 'TB.PCM.TollSKU', 'TB.PCM.Product.Manufacturer', 'TB.PCM.CabinetEnhancementsType', 'TB.PCM.CabinetEnchancementLocation', 'TB.PCM.Size', 'TB.PCM.ProductDescription', 'TB.PCM.DivisionSelected'],
    requiredFields: ['TB.PCM.Category', 'TB.PCM.ProductName', 'TB.PCM.Product.Manufacturer'],
    fallbacks: null
  }
};
// Dropdown items. A "group" scans multiple sub-sheets and consolidates them.
const DROPDOWN = [
  { key: 'tile', label: 'Tile', group: ['tile'] },
  { key: 'carpet', label: 'Carpet', group: ['carpet'] },
  { key: 'flooring', label: 'Flooring (Hardwood / LVP / Laminate)', group: ['flooring-hardwood', 'flooring-lvp', 'flooring-laminate'] },
  { key: 'cabinets', label: 'Cabinets (Vanities / Hardware / Enhancements)', group: ['cabinets-vanities', 'cabinets-hardware', 'cabinets-enhancements'] }
];

const CSS = `
  .g-wrap  { font-family: "Segoe UI", sans-serif; padding: 24px; max-width: 900px; }
  .g-title { font-size: 20px; font-weight: 600; margin-bottom: 2px; }
  .g-sub   { font-size: 13px; color: #555; margin-bottom: 18px; }
  .g-drop  { border: 2px dashed #aaa; border-radius: 8px; padding: 30px; text-align: center; cursor: pointer; color: #555; margin-bottom: 12px; }
  .g-drop.g-hover { border-color: #2b6cb0; background: #f0f6ff; color: #2b6cb0; }
  .g-row   { display: flex; gap: 10px; align-items: center; margin-bottom: 12px; flex-wrap: wrap; }
  .g-btn   { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; }
  .g-btn:disabled { opacity: .5; cursor: not-allowed; }
  .g-dry   { background: #edf2f7; color: #2d3748; }
  .g-go    { background: #2f855a; color: #fff; }
  .g-sync  { background: #6b46c1; color: #fff; }
  .g-sel   { padding: 6px; border: 1px solid #cbd5e0; border-radius: 6px; font-size: 13px; }
  .g-log   { background: #1a202c; color: #e2e8f0; font-family: monospace; font-size: 12px; padding: 14px; border-radius: 6px; margin-top: 14px; max-height: 360px; overflow: auto; white-space: pre-wrap; display: none; }
  .g-ok { color: #68d391; } .g-skip { color: #cbd5e0; } .g-err { color: #fc8181; } .g-info { color: #90cdf4; }
`;

const SHEETJS_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
function loadXLSX() {
  return new Promise((resolve, reject) => {
    if (window.XLSX) return resolve(window.XLSX);
    const s = document.createElement('script');
    s.src = SHEETJS_URL;
    s.onload = () => (window.XLSX ? resolve(window.XLSX) : reject(new Error('SheetJS loaded but window.XLSX missing.')));
    s.onerror = () => reject(new Error('Could not load SheetJS (check CSP / network).'));
    document.head.appendChild(s);
  });
}

function norm(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' '); }

function dsLabel(v, culture) {
  const L = v && (v.labels || v.Labels);
  if (!L) return '';
  if (typeof L.get === 'function') return L.get(culture) || L.get('en-US') || '';
  return L[culture] || L['en-US'] || '';
}

// Get one source's [[identifier, label]] live (cached), fall back to snapshot.
async function getSourcePairs(source, client, culture, log) {
  if (liveCache[source]) return liveCache[source];
  if (client && client.dataSources && typeof client.dataSources.getAsync === 'function') {
    try {
      const ds = await client.dataSources.getAsync(source);
      const values = (ds && (ds.values || ds.Values)) || [];
      const pairs = [];
      for (const v of values) { const id = v.identifier || v.Identifier; if (id) pairs.push([id, dsLabel(v, culture)]); }
      if (pairs.length) { liveCache[source] = pairs; return pairs; }
    } catch (e) { /* fall back */ }
  }
  const fb = DEFAULT_LOOKUPS[source] || [];
  liveCache[source] = fb;
  return fb;
}

function buildLabelMap(pairs) {
  const m = new Map();
  for (const [identifier, label] of pairs) {
    if (!identifier) continue;
    m.set(identifier.toLowerCase(), identifier);
    const last = identifier.split('.').pop();
    if (last) m.set(last.toLowerCase(), identifier);
    if (label) m.set(String(label).trim().toLowerCase(), identifier);
  }
  return m;
}

function resolveField(value, map) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return { value: '', bad: [] };
  const tokens = raw.split(/[|,]/).map(t => t.trim()).filter(Boolean);
  const ids = [], bad = [];
  for (const tok of tokens) {
    let id = map.get(tok.toLowerCase());
    if (!id && tok.indexOf('.') >= 0) id = tok;
    if (id) ids.push(id); else bad.push(tok);
  }
  return { value: ids.join('|'), bad };
}

async function parseFileAOA(file, XLSX, sheetName) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  let name = wb.SheetNames[0];
  if (sheetName) {
    const found = wb.SheetNames.find(n => String(n).trim().toLowerCase() === String(sheetName).trim().toLowerCase());
    if (found) name = found;
  }
  const ws = wb.Sheets[name];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  return { aoa, sheet: name, sheets: wb.SheetNames };
}

function ts() { const d = new Date(), p = n => String(n).padStart(2, '0'); return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`; }
function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// Build a CH-field record from one intake row (normalized-label -> value) for a config.
function buildRecord(cfg, rowObj) {
  const rec = {};
  if (cfg.supportsUpdate) {
    for (const [lbl, chf] of Object.entries(ID_LABELS)) {
      const v = (rowObj[lbl] || '').trim(); if (v && !rec[chf]) rec[chf] = v;
    }
  }
  if (cfg.itemCols) {
    const e1 = [], aoa = [];
    for (const it of cfg.itemCols) { const v = (rowObj[it.label] || '').trim(); if (v) { e1.push(v); aoa.push(it.aoa); } }
    if (e1.length) rec['TB.PCM.E1ItemNumber'] = e1.join(',');
    if (aoa.length) rec['TB.PCM.AreaOfApplication'] = aoa.join('|');
  }
  if (cfg.specialFeatures) {
    const ids = [];
    for (const f of cfg.specialFeatures.flags) { if ((rowObj[f.label] || '').trim()) f.ids.forEach(id => { if (!ids.includes(id)) ids.push(id); }); }
    if (ids.length) rec[cfg.specialFeatures.target] = ids.join('|');
  }
  for (const [lbl, chf] of Object.entries(cfg.fieldMap)) {
    const v = (rowObj[lbl] || '').trim(); if (v) rec[chf] = v;
  }
  if (cfg.fallbacks) {
    if (cfg.fallbacks.nameFrom) {
      let n = ''; for (const l of cfg.fallbacks.nameFrom) { const v = (rowObj[l] || '').trim(); if (v) { n = v; break; } }
      rec.__fallbackName = n;
    }
    if (cfg.fallbacks.skuFromItems) {
      rec.__fallbackSku = cfg.fallbacks.skuFromItems.map(l => (rowObj[l] || '').trim()).filter(Boolean).join(',');
    }
  }
  return rec;
}

// ---- Update matching -------------------------------------------------------
function keyFor(getVal, fields) {
  const parts = [];
  for (const f of fields) { const v = String(getVal(f) == null ? '' : getVal(f)).trim().toLowerCase(); if (!v) return null; parts.push(v); }
  return parts.join('|');
}
function buildExportIndex(aoa, strategies) {
  if (!aoa || aoa.length < 2) return { error: 'export has no data rows.' };
  const headers = aoa[0].map(h => String(h == null ? '' : h).trim().toLowerCase());
  const idi = headers.indexOf('id'), idnt = headers.indexOf('identifier');
  if (idi < 0 && idnt < 0) return { error: 'export must include id or identifier.' };
  const maps = strategies.map(() => new Map());
  for (let i = 1; i < aoa.length; i++) {
    const row = aoa[i];
    const values = {}; for (let j = 0; j < headers.length; j++) values[headers[j]] = row[j];
    const rowData = { id: idi >= 0 ? row[idi] : '', identifier: idnt >= 0 ? row[idnt] : '', values };
    strategies.forEach((fields, si) => {
      const key = keyFor(f => values[f.toLowerCase()], fields);
      if (key && !maps[si].has(key)) maps[si].set(key, rowData);
    });
  }
  return { maps };
}

// ---------------------------------------------------------------------------
export default function createExternalRoot(rootElement) {
  return {
    render(context) {
      const client = context && context.client;
      const culture = (context && (context.culture || (context.options && context.options.culture))) || 'en-US';

      const style = document.createElement('style'); style.textContent = CSS;
      const wrap = document.createElement('div'); wrap.className = 'g-wrap';
      wrap.innerHTML = `
        <div class="g-title">🏭 Product Import Generator</div>
        <div class="g-sub">Pick a category, upload its vendor <b>intake</b> spreadsheet, and download a
          ready-to-import <b>${SHEET_NAME}</b> workbook. Add <code>id</code>/<code>identifier</code> columns
          (plus a Content Hub export) for updates. Cabinets are create-only. Nothing is written to Content Hub.</div>
        <div class="g-row">
          <label style="font-size:13px;color:#555">Category:</label>
          <select id="g-cat" class="g-sel"></select>
          <span id="g-catnote" style="font-size:12px;color:#888"></span>
        </div>
        <div class="g-drop" id="g-drop">📎 <b>1. Intake file</b> — drop your vendor .xlsx / .csv here, or click to browse</div>
        <input type="file" id="g-file" accept=".xlsx,.xls,.csv" style="display:none" />
        <div class="g-drop" id="g-drop2">🔁 <b>2. Content Hub export</b> (optional — for updates) — drop the export with id/identifier</div>
        <input type="file" id="g-file2" accept=".xlsx,.xls,.csv" style="display:none" />
        <div class="g-row">
          <button class="g-btn g-dry" id="g-dry" disabled>🔍 Validate (dry run)</button>
          <button class="g-btn g-go"  id="g-go"  disabled>⬇ Generate import file</button>
          <span id="g-status" style="font-size:13px;color:#555"></span>
        </div>
        <div class="g-row">
          <button class="g-btn g-sync" id="g-sync">🔄 Reload option lists from Content Hub</button>
          <span id="g-lookups" style="font-size:12px;color:#888"></span>
        </div>
        <div class="g-log" id="g-log"></div>
      `;
      rootElement.innerHTML = ''; rootElement.appendChild(style); rootElement.appendChild(wrap);

      const catSel = wrap.querySelector('#g-cat'), catNote = wrap.querySelector('#g-catnote');
      const drop = wrap.querySelector('#g-drop'), input = wrap.querySelector('#g-file');
      const drop2 = wrap.querySelector('#g-drop2'), input2 = wrap.querySelector('#g-file2');
      const dryBtn = wrap.querySelector('#g-dry'), goBtn = wrap.querySelector('#g-go');
      const syncBtn = wrap.querySelector('#g-sync'), lookupsStatus = wrap.querySelector('#g-lookups');
      const status = wrap.querySelector('#g-status'), logEl = wrap.querySelector('#g-log');
      let currentFile = null, currentExport = null;

      catSel.innerHTML = DROPDOWN.map(d => `<option value="${d.key}">${d.label}</option>`).join('');
      function partsForSelection() { const item = DROPDOWN.find(d => d.key === catSel.value) || DROPDOWN[0]; return { item, parts: item.group.map(k => CATEGORY_CONFIGS[k]) }; }
      function updateCatNote() {
        const { parts } = partsForSelection();
        const upd = parts.every(p => p.supportsUpdate);
        catNote.textContent = (parts.length > 1 ? `consolidates ${parts.length} sheets · ` : '') + (upd ? 'create + update' : 'create-only');
      }
      updateCatNote();
      catSel.addEventListener('change', updateCatNote);

      const lookupCount = () => Object.values(liveCache).reduce((a, v) => a + (v ? v.length : 0), 0);
      const setLookupsStatus = note => { lookupsStatus.textContent = `Option lists: ${lookupCount()} values ${note}`; };
      setLookupsStatus('— read live from Content Hub on first run.');

      function log(msg, cls) {
        logEl.style.display = 'block';
        const line = document.createElement('div'); if (cls) line.className = cls;
        line.textContent = msg; logEl.appendChild(line); logEl.scrollTop = logEl.scrollHeight;
      }
      function clearLog() { logEl.innerHTML = ''; logEl.style.display = 'none'; }

      function refreshStatus() {
        const parts = [];
        if (currentFile) parts.push(`intake: ${currentFile.name}`);
        if (currentExport) parts.push(`export: ${currentExport.name}`);
        status.textContent = parts.join('  ·  ');
        dryBtn.disabled = !currentFile; goBtn.disabled = !currentFile;
      }
      const onFile = f => { if (!f) return; clearLog(); currentFile = f; refreshStatus(); };
      const onExport = f => { if (!f) return; clearLog(); currentExport = f; refreshStatus(); };

      drop.addEventListener('click', () => input.click());
      input.addEventListener('change', e => onFile(e.target.files[0]));
      drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('g-hover'); });
      drop.addEventListener('dragleave', () => drop.classList.remove('g-hover'));
      drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('g-hover'); onFile(e.dataTransfer.files[0]); });
      drop2.addEventListener('click', () => input2.click());
      input2.addEventListener('change', e => onExport(e.target.files[0]));
      drop2.addEventListener('dragover', e => { e.preventDefault(); drop2.classList.add('g-hover'); });
      drop2.addEventListener('dragleave', () => drop2.classList.remove('g-hover'));
      drop2.addEventListener('drop', e => { e.preventDefault(); drop2.classList.remove('g-hover'); onExport(e.dataTransfer.files[0]); });

      async function run(dryRun) {
        clearLog(); dryBtn.disabled = true; goBtn.disabled = true;
        const { item, parts } = partsForSelection();
        const groupSupportsUpdate = parts.every(p => p.supportsUpdate);
        log(`${dryRun ? '── VALIDATE' : '── GENERATE'} · ${item.label} ──`, 'g-info');
        try {
          const XLSX = await loadXLSX();
          const wb = XLSX.read(await currentFile.arrayBuffer(), { type: 'array' });
          const sheetNames = wb.SheetNames;

          const allRecords = [];
          const optionFields = new Set();
          const outColsOrder = [];
          const usedSheets = new Set();

          for (const part of parts) {
            // Pick this part's sheet (by keyword match; single categories use the first sheet).
            let name = null;
            if (part.sheetMatch) {
              for (const kw of part.sheetMatch) { const f = sheetNames.find(n => !usedSheets.has(n) && norm(n).includes(kw)); if (f) { name = f; break; } }
              if (!name) { if (parts.length > 1) log(`No sheet for ${part.label} (looking for: ${part.sheetMatch.join(', ')}).`, 'g-skip'); continue; }
            } else { name = sheetNames[0]; }
            usedSheets.add(name);
            const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: false });
            if (!aoa || aoa.length < 2) { log(`Sheet "${name}": no data — skipped.`, 'g-skip'); continue; }

            const rawHeaders = aoa[0].map(h => String(h == null ? '' : h).trim());
            const normHeaders = rawHeaders.map(norm);
            const known = new Set([...Object.keys(part.fieldMap), ...Object.keys(ID_LABELS)]);
            if (part.itemCols) part.itemCols.forEach(i => known.add(i.label));
            if (part.specialFeatures) part.specialFeatures.flags.forEach(f => known.add(f.label));
            if (part.fallbacks && part.fallbacks.nameFrom) part.fallbacks.nameFrom.forEach(l => known.add(l));

            let n = 0;
            for (let i = 1; i < aoa.length; i++) {
              const row = aoa[i];
              if (!row.some(c => String(c == null ? '' : c).trim() !== '')) continue;
              const obj = {}; for (let j = 0; j < normHeaders.length; j++) if (normHeaders[j]) obj[normHeaders[j]] = row[j];
              const rec = buildRecord(part, obj);
              rec.__row = i + 1; rec.__cfg = part; rec.__sheet = name;
              if (part.categoryValue) rec['TB.PCM.Category'] = part.categoryValue;
              allRecords.push(rec); n++;
            }
            const skipped = rawHeaders.filter((h, i) => h && !known.has(normHeaders[i]));
            log(`Sheet "${name}" → ${part.categoryValue ? part.categoryValue.split('.').pop() : ''}: ${n} row(s).`, 'g-info');
            if (skipped.length) log(`  skipped: ${skipped.join(', ')}`, 'g-skip');
            part.optionListFields.forEach(f => optionFields.add(f));
            part.outCols.forEach(c => { if (!outColsOrder.includes(c)) outColsOrder.push(c); });
          }

          if (!allRecords.length) { log('No data rows found in any sheet.', 'g-err'); return; }
          if (parts.length > 1) log(`Total: ${allRecords.length} row(s) across ${usedSheets.size} sheet(s).`, 'g-info');

          // Update matching (only if the whole selection supports it).
          const wantUpdate = !!currentExport;
          const updateMode = wantUpdate && groupSupportsUpdate;
          if (wantUpdate && !groupSupportsUpdate) log('This category is create-only — the Content Hub export is ignored.', 'g-skip');
          let outputRecords = allRecords;
          if (updateMode) {
            const strategies = parts[0].matchStrategies;
            const expWb = XLSX.read(await currentExport.arrayBuffer(), { type: 'array' });
            const expAoa = XLSX.utils.sheet_to_json(expWb.Sheets[expWb.SheetNames[0]], { header: 1, defval: '', raw: false });
            const { maps, error } = buildExportIndex(expAoa, strategies);
            if (error) { log(`✗ Content Hub export: ${error}`, 'g-err'); return; }
            log(`UPDATE mode — matching by: ${strategies.map(s => s.join('+')).join('  or  ')}`, 'g-info');
            let matched = 0; const unmatched = [];
            for (const r of allRecords) {
              let hit = null;
              for (let si = 0; si < strategies.length; si++) { const key = keyFor(f => r[f], strategies[si]); if (key && maps[si].has(key)) { hit = maps[si].get(key); break; } }
              if (hit) {
                if (hit.id != null && String(hit.id).trim()) r['id'] = hit.id;
                if (hit.identifier != null && String(hit.identifier).trim()) r['identifier'] = hit.identifier;
                r.__matched = true; r.__export = hit.values; matched++;
              } else { r.__matched = false; unmatched.push(r.__row); }
            }
            log(`Matched ${matched} of ${allRecords.length} row(s).`, matched ? 'g-ok' : 'g-err');
            if (unmatched.length) log(`Unmatched (skipped): row ${unmatched.join(', ')}`, 'g-err');
            outputRecords = allRecords.filter(r => r.__matched);
            if (!outputRecords.length) { log('Nothing to update.', 'g-err'); return; }
          }

          // NEW records only: per-record fallbacks for ProductName/SKU.
          if (!updateMode) {
            for (const r of outputRecords) {
              if (!String(r['TB.PCM.ProductName'] == null ? '' : r['TB.PCM.ProductName']).trim() && r.__fallbackName) r['TB.PCM.ProductName'] = r.__fallbackName;
              if (!String(r['TB.PCM.Product.SKU'] == null ? '' : r['TB.PCM.Product.SKU']).trim() && r.__fallbackSku) r['TB.PCM.Product.SKU'] = r.__fallbackSku;
            }
          }

          const used = new Set();
          outputRecords.forEach(r => Object.keys(r).forEach(k => { if (!k.startsWith('__')) used.add(k); }));

          // Resolve option-list fields (union across the processed sheets).
          const meta = {};
          for (const f of optionFields) {
            if (!used.has(f)) continue;
            const src = sourceFor(f);
            const pairs = await getSourcePairs(src, client, culture, log);
            meta[f] = { isOption: pairs.length > 0, map: buildLabelMap(pairs) };
            if (meta[f].isOption) log(`  ${f} → ${src}: ${pairs.length} options`, 'g-info');
            else log(`  ${f}: no option list (${src}) — passed through as text`, 'g-skip');
          }
          const unresolved = [];
          for (const r of outputRecords) {
            for (const f of Object.keys(r)) {
              if (f.startsWith('__')) continue;
              if (meta[f] && meta[f].isOption) {
                const res = resolveField(r[f], meta[f].map);
                res.bad.forEach(b => unresolved.push({ row: r.__row, field: f, value: b }));
                if (res.bad.length && updateMode) { const o = r.__export ? r.__export[f.toLowerCase()] : ''; r[f] = (o == null ? '' : o); }
                else r[f] = res.value;
              }
            }
          }
          if (unresolved.length === 0) log('All option-list values matched. ✓', 'g-ok');
          else if (updateMode) { log(`⚠ ${unresolved.length} value(s) not found — existing value kept:`, 'g-err'); unresolved.slice(0, 100).forEach(u => log(`   Row ${u.row} · ${u.field}: "${u.value}" (kept)`, 'g-err')); }
          else { log(`⚠ ${unresolved.length} value(s) not found — left BLANK:`, 'g-err'); unresolved.slice(0, 100).forEach(u => log(`   Row ${u.row} · ${u.field}: "${u.value}"`, 'g-err')); }
          if (unresolved.length > 100) log(`   … and ${unresolved.length - 100} more`, 'g-err');

          // Required-field check per record's sub-config (new records only).
          if (!updateMode) {
            const missing = [];
            for (const r of outputRecords) { const m = r.__cfg.requiredFields.filter(f => !String(r[f] == null ? '' : r[f]).trim()); if (m.length) missing.push({ row: r.__row, m }); }
            if (missing.length) { log(`⚠ Required field(s) missing on ${missing.length} new record(s):`, 'g-err'); missing.slice(0, 100).forEach(x => log(`   Row ${x.row}: ${x.m.join(', ')}`, 'g-err')); }
            else log('All required fields present. ✓', 'g-ok');
          }

          if (dryRun) { log('Dry run complete — review above, then Generate.', 'g-info'); return; }

          const outCols = outColsOrder.filter(f => used.has(f));
          const outRows = [outCols];
          for (const r of outputRecords) outRows.push(outCols.map(f => (r[f] == null ? '' : r[f])));
          const ws = XLSX.utils.aoa_to_sheet(outRows);
          const wbOut = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wbOut, ws, SHEET_NAME);
          const arr = XLSX.write(wbOut, { bookType: 'xlsx', type: 'array' });
          const tag = item.label.replace(/[^a-z0-9]+/gi, '');
          const fname = `ContentHub_${tag}Import_${ts()}.xlsx`;
          downloadBlob(new Blob([arr], { type: 'application/octet-stream' }), fname);
          log(`✓ Generated ${fname} — ${outputRecords.length} ${updateMode ? 'update' : 'new'} row(s).`, 'g-ok');
        } catch (e) {
          log(`✗ ${e && e.message ? e.message : e}`, 'g-err');
        } finally { dryBtn.disabled = false; goBtn.disabled = false; }
      }

      dryBtn.addEventListener('click', () => run(true));
      goBtn.addEventListener('click', () => run(false));

      syncBtn.addEventListener('click', async () => {
        clearLog(); syncBtn.disabled = true;
        log('── RELOAD OPTION LISTS FROM CONTENT HUB ──', 'g-info');
        liveCache = {};
        let total = 0;
        for (const name of DATA_SOURCE_NAMES) {
          const pairs = await getSourcePairs(name, client, culture, log);
          total += pairs.length;
          log(`  ${name}: ${pairs.length} value(s)`, pairs.length ? 'g-ok' : 'g-err');
        }
        setLookupsStatus('— reloaded live from Content Hub.');
        log(`✓ Loaded ${total} value(s) across ${DATA_SOURCE_NAMES.length} data source(s). (Divisions & Category use the snapshot.)`, 'g-ok');
        syncBtn.disabled = false;
      });
    },
    unmount() { rootElement.innerHTML = ''; }
  };
}
