function createShopifySheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = ss.getSheetByName("Sheet2");

  if (!target) {
    throw new Error("Sheet2 was not found.");
  }

  // =====================================================
  // XML FEED
  // =====================================================

  const XML_URL =
    "https://platinumhellas.gr/wp-content/uploads/woo-product-feed-pro/xml/MlEnhpqlfasKVYo7kLAnZfx9GoAXpWte.xml";

  // =====================================================
  // DOWNLOAD XML
  // =====================================================

  const response = UrlFetchApp.fetch(XML_URL, {
    muteHttpExceptions: true,
    followRedirects: true
  });

  const status = response.getResponseCode();

  if (status !== 200) {
    throw new Error(
      "Could not download XML feed. HTTP status: " + status
    );
  }

  const xmlText = response.getContentText();

  if (!xmlText || xmlText.trim() === "") {
    throw new Error("The XML feed is empty.");
  }

  // =====================================================
  // PARSE XML
  // =====================================================

  let document;

  try {

    document =
      XmlService.parse(xmlText);

  } catch (error) {

    throw new Error(
      "Could not parse XML feed: " +
      error.message
    );

  }

  const root =
    document.getRootElement();

  // =====================================================
  // FIND ITEMS
  // =====================================================

  const items = [];

  collectItems(
    root,
    items
  );

  if (items.length === 0) {

    throw new Error(
      "No <item> elements were found in the XML feed."
    );

  }

  Logger.log(
    "XML items found: " +
    items.length
  );

  // =====================================================
  // SHOPIFY HEADERS
  // =====================================================

  const headers = [

    "Title",
    "URL handle",
    "Description",
    "Vendor",
    "Product category",
    "Type",
    "Tags",
    "Published on online store",
    "Status",

    "SKU",
    "Barcode",

    "Option1 name",
    "Option1 value",
    "Option1 Linked To",

    "Option2 name",
    "Option2 value",
    "Option2 Linked To",

    "Option3 name",
    "Option3 value",
    "Option3 Linked To",

    "Price",
    "Compare-at price",
    "Cost per item",

    "Charge tax",
    "Tax code",

    "Unit price total measure",
    "Unit price total measure unit",
    "Unit price base measure",
    "Unit price base measure unit",

    "Inventory tracker",
    "Inventory quantity",
    "Continue selling when out of stock",

    "Weight value (grams)",
    "Weight unit for display",

    "Requires shipping",
    "Fulfillment service",

    "Product image URL",
    "Image position",
    "Image alt text",
    "Variant image URL",

    "Gift card",

    "SEO title",
    "SEO description",

    "Color (product.metafields.shopify.color-pattern)",

    "Google Shopping / Google product category",
    "Google Shopping / Gender",
    "Google Shopping / Age group",
    "Google Shopping / Manufacturer part number (MPN)",
    "Google Shopping / Ad group name",
    "Google Shopping / Ads labels",
    "Google Shopping / Condition",
    "Google Shopping / Custom product",

    "Google Shopping / Custom label 0",
    "Google Shopping / Custom label 1",
    "Google Shopping / Custom label 2",
    "Google Shopping / Custom label 3",
    "Google Shopping / Custom label 4"

  ];

  const output = [headers];

  // =====================================================
  // GROUP PRODUCTS
  // =====================================================

  const groups = {};

  items.forEach(function(item) {

    const id =
      getXmlValue(
        item,
        "id"
      );

    const groupId =
      getXmlValue(
        item,
        "item_group_id"
      );

    if (!id) {
      return;
    }

    const productGroup =
      groupId || id;

    if (!groups[productGroup]) {
      groups[productGroup] = [];
    }

    groups[productGroup].push(item);

  });

  // =====================================================
  // PROCESS PRODUCT GROUPS
  // =====================================================

  Object.keys(groups).forEach(function(groupId) {

    const variants =
      groups[groupId];

    if (!variants.length) {
      return;
    }

    const first =
      variants[0];

    // ===================================================
    // PRODUCT LEVEL DATA
    // ===================================================

    const firstTitle =
      getXmlValue(
        first,
        "title"
      );

    const productTitle =
      removePackagingFromTitle(
        firstTitle
      );

    const productDescription =
      getXmlValue(
        first,
        "description"
      );

    const productBrand =
      getXmlValue(
        first,
        "brand"
      );

    const productType =
      getXmlValue(
        first,
        "product_type"
      );

    const productGoogleCategory =
      getXmlValue(
        first,
        "google_product_category"
      );

    const productCondition =
      getXmlValue(
        first,
        "condition"
      );

    // ===================================================
    // PRODUCT IMAGE
    // ===================================================

    const productImage =
      getBestImageUrl(
        first
      );

    // ===================================================
    // MULTIPLE VARIANTS
    // ===================================================

    const hasMultipleVariants =
      variants.length > 1;

    // ===================================================
    // HANDLE
    // ===================================================

    const firstLink =
      getXmlValue(
        first,
        "link"
      );

    let groupHandle =
      createBaseHandleFromUrl(
        firstLink
      );

    // ---------------------------------------------------
    // Fallback handle from title
    // ---------------------------------------------------

    if (!groupHandle) {

      groupHandle =
        createShopifyHandle(
          productTitle
        );

    }

    Logger.log(
      "=========================================="
    );

    Logger.log(
      "GROUP: " +
      groupId
    );

    Logger.log(
      "Variants: " +
      variants.length
    );

    Logger.log(
      "Product title: " +
      productTitle
    );

    Logger.log(
      "Handle: " +
      groupHandle
    );

    Logger.log(
      "Product image: " +
      productImage
    );

    // ===================================================
    // EACH VARIANT
    // ===================================================

    variants.forEach(function(item) {

      // -------------------------------------------------
      // BASIC DATA
      // -------------------------------------------------

      const id =
        getXmlValue(
          item,
          "id"
        );

      const originalTitle =
        getXmlValue(
          item,
          "title"
        );

      const title =
        removePackagingFromTitle(
          originalTitle
        );

      const description =
        getXmlValue(
          item,
          "description"
        );

      const brand =
        getXmlValue(
          item,
          "brand"
        );

      const type =
        getXmlValue(
          item,
          "product_type"
        );

      const link =
        getXmlValue(
          item,
          "link"
        );

      // -------------------------------------------------
      // IMAGE
      // -------------------------------------------------

      let image =
        getBestImageUrl(
          item
        );

      // -------------------------------------------------
      // If variant image missing, use product image
      // -------------------------------------------------

      if (!image) {
        image = productImage;
      }

      const gtin =
        getXmlValue(
          item,
          "gtin"
        );

      const googleCategory =
        getXmlValue(
          item,
          "google_product_category"
        );

      const condition =
        getXmlValue(
          item,
          "condition"
        );

      // =================================================
      // PRICE
      // =================================================

      const originalPriceRaw =
        getXmlValue(
          item,
          "price"
        );

      const salePriceRaw =
        getXmlValue(
          item,
          "sale_price"
        );

      const originalPrice =
        parsePrice(
          originalPriceRaw
        );

      const salePrice =
        parsePrice(
          salePriceRaw
        );

      let shopifyPrice = "";
      let compareAtPrice = "";

      // -------------------------------------------------
      // Valid sale
      // -------------------------------------------------

      if (
        originalPrice !== "" &&
        salePrice !== "" &&
        Number(salePrice) <
        Number(originalPrice)
      ) {

        shopifyPrice =
          salePrice;

        compareAtPrice =
          originalPrice;

      }

      // -------------------------------------------------
      // Original price only
      // -------------------------------------------------

      else if (
        originalPrice !== ""
      ) {

        shopifyPrice =
          originalPrice;

      }

      // -------------------------------------------------
      // Sale price only
      // -------------------------------------------------

      else if (
        salePrice !== ""
      ) {

        shopifyPrice =
          salePrice;

      }

      // =================================================
      // PACKAGING
      // =================================================

      const rawPackagingValue =
        getPackagingValueFromUrl(
          link
        );

      const formattedPackagingValue =
        formatPackagingValue(
          rawPackagingValue
        );

      // =================================================
      // OPTION
      // =================================================

      let option1Name = "";
      let option1Value = "";

      if (hasMultipleVariants) {

        option1Name =
          "Συσκευασία";

        if (formattedPackagingValue) {

          option1Value =
            formattedPackagingValue;

        } else {

          option1Value =
            "Variant " + id;

        }

      } else {

        option1Name =
          "Title";

        option1Value =
          "Default Title";

      }

      // =================================================
      // SEO
      // =================================================

      const seoTitle =
        productTitle;

      const seoDescription =
        stripHTML(
          description ||
          productDescription
        );

      // =================================================
      // IMAGE FIELDS
      // =================================================

      let imageUrl = "";
      let imagePosition = "";
      let imageAlt = "";
      let variantImageUrl = "";

      if (image) {

        imageUrl =
          image;

        imagePosition =
          "1";

        imageAlt =
          title ||
          productTitle;

        variantImageUrl =
          image;

      }

      // =================================================
      // SHOPIFY ROW
      // =================================================

      const row = [

        // 1 Title
        productTitle,

        // 2 URL handle
        groupHandle,

        // 3 Description
        description ||
        productDescription,

        // 4 Vendor
        brand ||
        productBrand,

        // 5 Product category
        "",

        // 6 Type
        type ||
        productType,

        // 7 Tags
        "",

        // 8 Published
        "TRUE",

        // 9 Status
        "ACTIVE",

        // 10 SKU
        id,

        // 11 Barcode
        gtin,

        // 12 Option1 name
        option1Name,

        // 13 Option1 value
        option1Value,

        // 14 Option1 Linked To
        "",

        // 15 Option2 name
        "",

        // 16 Option2 value
        "",

        // 17 Option2 Linked To
        "",

        // 18 Option3 name
        "",

        // 19 Option3 value
        "",

        // 20 Option3 Linked To
        "",

        // 21 Price
        shopifyPrice,

        // 22 Compare-at price
        compareAtPrice,

        // 23 Cost per item
        "",

        // 24 Charge tax
        "TRUE",

        // 25 Tax code
        "",

        // 26 Unit price total measure
        "",

        // 27 Unit price total measure unit
        "",

        // 28 Unit price base measure
        "",

        // 29 Unit price base measure unit
        "",

        // 30 Inventory tracker
        "shopify",

        // 31 Inventory quantity
        100,

        // 32 Continue selling
        "deny",

        // 33 Weight value
        "",

        // 34 Weight unit
        "g",

        // 35 Requires shipping
        "TRUE",

        // 36 Fulfillment service
        "manual",

        // 37 Product image URL
        imageUrl,

        // 38 Image position
        imagePosition,

        // 39 Image alt text
        imageAlt,

        // 40 Variant image URL
        variantImageUrl,

        // 41 Gift card
        "FALSE",

        // 42 SEO title
        seoTitle,

        // 43 SEO description
        seoDescription,

        // 44 Color
        "",

        // 45 Google category
        googleCategory ||
        productGoogleCategory,

        // 46 Gender
        "",

        // 47 Age group
        "",

        // 48 MPN
        id,

        // 49 Ad group
        "",

        // 50 Ads labels
        "",

        // 51 Condition
        condition ||
        productCondition ||
        "new",

        // 52 Custom product
        "",

        // 53 Custom label 0
        "",

        // 54 Custom label 1
        "",

        // 55 Custom label 2
        "",

        // 56 Custom label 3
        "",

        // 57 Custom label 4
        ""

      ];

      output.push(row);

      // =================================================
      // DEBUG
      // =================================================

      Logger.log(
        "SKU: " +
        id +
        " | Original title: " +
        originalTitle +
        " | Final title: " +
        productTitle +
        " | Raw packaging: " +
        rawPackagingValue +
        " | Variant: " +
        option1Value +
        " | Image: " +
        imageUrl
      );

    });

  });

  // =====================================================
  // WRITE SHEET
  // =====================================================

  target.clearContents();

  target
    .getRange(
      1,
      1,
      output.length,
      headers.length
    )
    .setValues(output);

  // =====================================================
  // FORMAT
  // =====================================================

  target.setFrozenRows(1);

  target
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setFontWeight("bold");

  if (output.length > 1) {

    target
      .getRange(
        2,
        21,
        output.length - 1,
        2
      )
      .setNumberFormat("0.00");

    target
      .getRange(
        2,
        31,
        output.length - 1,
        1
      )
      .setNumberFormat("0");

  }

  SpreadsheetApp.flush();

  // =====================================================
  // SUMMARY
  // =====================================================

  let rowsWithImages = 0;
  let rowsWithoutImages = 0;

  for (
    let r = 1;
    r < output.length;
    r++
  ) {

    const image =
      output[r][36];

    if (image) {
      rowsWithImages++;
    } else {
      rowsWithoutImages++;
    }

  }

  Logger.log(
    "=========================================="
  );

  Logger.log(
    "XML items: " +
    items.length
  );

  Logger.log(
    "Product groups: " +
    Object.keys(groups).length
  );

  Logger.log(
    "Shopify rows: " +
    (output.length - 1)
  );

  Logger.log(
    "Rows with images: " +
    rowsWithImages
  );

  Logger.log(
    "Rows without images: " +
    rowsWithoutImages
  );

  Logger.log(
    "=========================================="
  );

}


// =====================================================
// REMOVE PACKAGING FROM PRODUCT TITLE
// =====================================================
//
// Examples
//
// PLATINUM PUPPY ... 30kg (6 x 5 kg)
// → PLATINUM PUPPY ...
//
// PLATINUM PUPPY ... 5kg
// → PLATINUM PUPPY ...
//
// PLATINUM PUPPY ... 1.5kg
// → PLATINUM PUPPY ...
//
// PLATINUM ... 6x185gr
// → PLATINUM ...
//
// PLATINUM ... 6 x 185 gr
// → PLATINUM ...
//
// The function only removes packaging when it appears
// at the END of the title.
// =====================================================

function removePackagingFromTitle(title) {

  if (!title) {
    return "";
  }

  let value =
    String(title)
      .trim();

  // =====================================================
  // Repeat a few times
  //
  // This handles titles such as
  //
  // Product 30kg (6 x 5 kg)
  //
  // by first removing the bracket and then the 30kg.
  // =====================================================

  for (
    let i = 0;
    i < 5;
    i++
  ) {

    const before =
      value;

    // ---------------------------------------------------
    // Remove bracketed multipack
    //
    // (6 x 5 kg)
    // (6x5kg)
    // (3 x 1.5 kg)
    // (2 x 500 gr)
    // ---------------------------------------------------

    value =
      value.replace(
        /\s*\(\s*\d+\s*[x×]\s*\d+(?:[.,]\d+)?\s*(?:kg|kgr|g|gr|ml|l)\s*\)\s*$/i,
        ""
      );

    // ---------------------------------------------------
    // Remove bracketed decimal with dash notation
    //
    // (3-1-5-kg)
    // ---------------------------------------------------

    value =
      value.replace(
        /\s*\(\s*\d+(?:-\d+){1,2}-(?:kg|kgr|g|gr|ml|l)\s*\)\s*$/i,
        ""
      );

    // ---------------------------------------------------
    // Remove bracketed whole weight
    //
    // (30kg)
    // (30 kg)
    // ---------------------------------------------------

    value =
      value.replace(
        /\s*\(\s*\d+(?:[.,]\d+)?\s*(?:kg|kgr|g|gr|ml|l)\s*\)\s*$/i,
        ""
      );

    // ---------------------------------------------------
    // Remove normal multipack
    //
    // 6 x 5 kg
    // 6x5kg
    // 3 x 1.5 kg
    // ---------------------------------------------------

    value =
      value.replace(
        /\s+\d+\s*[x×]\s*\d+(?:[.,]\d+)?\s*(?:kg|kgr|g|gr|ml|l)\s*$/i,
        ""
      );

    // ---------------------------------------------------
    // Remove dash multipack
    //
    // 6-x-5-kg
    // 3-x-1-5-kg
    // ---------------------------------------------------

    value =
      value.replace(
        /\s+\d+-x-\d+(?:-\d+)?-(?:kg|kgr|g|gr|ml|l)\s*$/i,
        ""
      );

    // ---------------------------------------------------
    // Remove compact multipack
    //
    // 6x185gr
    // 3x1.5kg
    // ---------------------------------------------------

    value =
      value.replace(
        /\s+\d+x\d+(?:[.,]\d+)?\s*(?:kg|kgr|g|gr|ml|l)\s*$/i,
        ""
      );

    // ---------------------------------------------------
    // Remove WooCommerce decimal
    //
    // 1-5-kg
    // ---------------------------------------------------

    value =
      value.replace(
        /\s+\d+-\d+-(?:kg|kgr|g|gr|ml|l)\s*$/i,
        ""
      );

    // ---------------------------------------------------
    // Remove standard decimal
    //
    // 1.5kg
    // 1.5 kg
    // ---------------------------------------------------

    value =
      value.replace(
        /\s+\d+[.,]\d+\s*(?:kg|kgr|g|gr|ml|l)\s*$/i,
        ""
      );

    // ---------------------------------------------------
    // Remove standard whole weight
    //
    // 30kg
    // 30 kg
    // 500gr
    // 500 gr
    // ---------------------------------------------------

    value =
      value.replace(
        /\s+\d+\s*(?:kg|kgr|g|gr|ml|l)\s*$/i,
        ""
      );

    // ---------------------------------------------------
    // Stop when nothing changed
    // ---------------------------------------------------

    if (
      value === before
    ) {
      break;
    }

  }

  // =====================================================
  // Clean trailing separators
  // =====================================================

  value =
    value.replace(
      /[\s\-–—|]+$/g,
      ""
    );

  return value.trim();

}


// =====================================================
// GET PACKAGING VALUE FROM URL
// =====================================================

function getPackagingValueFromUrl(url) {

  if (!url) {
    return "";
  }

  const questionMark =
    url.indexOf("?");

  if (questionMark === -1) {
    return "";
  }

  const query =
    url.substring(
      questionMark + 1
    );

  const params =
    query.split("&");

  // =====================================================
  // Exact packaging attribute
  // =====================================================

  for (
    let i = 0;
    i < params.length;
    i++
  ) {

    const parts =
      params[i].split("=");

    if (parts.length < 2) {
      continue;
    }

    let key =
      parts[0];

    let value =
      parts
        .slice(1)
        .join("=");

    try {

      key =
        decodeURIComponent(key);

      value =
        decodeURIComponent(value);

    } catch (e) {}

    if (
      key
        .toLowerCase()
        .trim() ===
      "attribute_pa_packaging-size"
    ) {

      return value.trim();

    }

  }

  // =====================================================
  // Fallback to any WooCommerce attribute
  // =====================================================

  for (
    let i = 0;
    i < params.length;
    i++
  ) {

    const parts =
      params[i].split("=");

    if (parts.length < 2) {
      continue;
    }

    let key =
      parts[0];

    let value =
      parts
        .slice(1)
        .join("=");

    try {

      key =
        decodeURIComponent(key);

      value =
        decodeURIComponent(value);

    } catch (e) {}

    if (
      key
        .toLowerCase()
        .indexOf("attribute_pa_") === 0
    ) {

      return value.trim();

    }

  }

  return "";

}


// =====================================================
// FORMAT PACKAGING VALUE
// =====================================================

// =====================================================
// EXTRACT PACKAGING FROM TITLE (FIXED)
// =====================================================
function extractPackagingFromTitle(title) {
  if (!title) return "";
  let value = String(title).trim().replace(/×/g, "x");

  // 1. Parenthesized multipack (6 x 5 kg) ή (2x900gr)
  let match = value.match(/\(\s*(\d+)\s*x\s*(\d+(?:[.,]\d+)?)\s*(kg|kgr|g|gr|ml|l)\s*\)\s*$/i);
  if (match) {
    let unit = match[3].toLowerCase() === "kgr" ? "kg" : match[3].toLowerCase();
    return match[1] + "x" + match[2].replace(",", ".") + unit;
  }

  // 2. Multipack με x (2x900g ή 2 x 900 g)
  match = value.match(/(\d+)\s*x\s*(\d+(?:[.,]\d+)?)\s*(kg|kgr|g|gr|ml|l)\s*$/i);
  if (match) {
    let unit = match[3].toLowerCase() === "kgr" ? "kg" : match[3].toLowerCase();
    return match[1] + "x" + match[2].replace(",", ".") + unit;
  }

  // 3. Multipack με παύλα στον τίτλο (2-900g ή 2-900-g)
  match = value.match(/(\d+)\s*-\s*(\d+)\s*-?\s*(kg|kgr|g|gr|ml|l)\s*$/i);
  if (match) {
    let unit = match[3].toLowerCase() === "kgr" ? "kg" : match[3].toLowerCase();
    return match[1] + "x" + match[2] + unit;
  }

  // 4. Δεκαδικά (1.5kg ή 1,5 kg)
  match = value.match(/(\d+[.,]\d+)\s*(kg|kgr|g|gr|ml|l)\s*$/i);
  if (match) {
    let unit = match[2].toLowerCase() === "kgr" ? "kg" : match[2].toLowerCase();
    return match[1].replace(",", ".") + unit;
  }

  // 5. Απλό βάρος (10kg ή 500g)
  match = value.match(/(\d+)\s*(kg|kgr|g|gr|ml|l)\s*$/i);
  if (match) {
    let unit = match[2].toLowerCase() === "kgr" ? "kg" : match[2].toLowerCase();
    return match[1] + unit;
  }

  return "";
}


// =====================================================
// FORMAT PACKAGING VALUE FROM URL (FIXED)
// =====================================================
function formatPackagingValue(rawValue) {
  if (!rawValue) return "";
  let value = String(rawValue).trim().toLowerCase();
  try { value = decodeURIComponent(value); } catch (e) {}
  value = value.replace(/×/g, "x").replace(/\s*x\s*/gi, "x");


  // 2. Μορφή 2-x-900-g
  match = value.match(/^(\d+)-x-(\d+)-(kg|kgr|g|gr|ml|l)$/i);
  if (match) {
    let unit = match[3].toLowerCase() === "kgr" ? "kg" : match[3].toLowerCase();
    return match[1] + "x" + match[2] + unit;
  }

  // 3. Κανονικό x (2x900g)
  match = value.match(/^(\d+)x(\d+)\s*(kg|kgr|g|gr|ml|l)$/i);
  if (match) {
    let unit = match[3].toLowerCase() === "kgr" ? "kg" : match[3].toLowerCase();
    return match[1] + "x" + match[2] + unit;
  }



    // =====================================================
  // 10-kg
  //
  // → 10kg
  // =====================================================

  match =
    value.match(
      /^(\d+)-(kg|kgr|g|gr|ml|l)$/i
    );

  if (match) {

    return (
      match[1] +
      normalizeUnit(match[2])
    );

  }

  // =====================================================
  // 1.5kg
  // =====================================================

  match =
    value.match(
      /^(\d+)\.(\d+)\s*(kg|kgr|g|gr|ml|l)$/i
    );

  if (match) {

    return (
      match[1] +
      "." +
      match[2] +
      normalizeUnit(match[3])
    );

  }

  // =====================================================
  // 10kg
  // =====================================================

  match =
    value.match(
      /^(\d+)\s*(kg|kgr|g|gr|ml|l)$/i
    );

  if (match) {

    return (
      match[1] +
      normalizeUnit(match[2])
    );

  }

  return value;

}


// =====================================================
// NORMALIZE UNIT
// =====================================================

function normalizeUnit(unit) {
  if (!unit) return "";
  let u = unit.toLowerCase().trim();
  if (u === "kgr") return "kg";
  if (u === "gr") return "g";
  return u;
}


// =====================================================
// GET BEST IMAGE URL
// =====================================================
//
// WooCommerce / Google feeds can use different image
// fields. We check several possible fields.
//
// Priority
//
// image_link
// image
// featured_image
// main_image
// product_image
// additional_image_link
//
// =====================================================

function getBestImageUrl(item) {

  if (!item) {
    return "";
  }

  const possibleFields = [

    "image_link",
    "image",
    "featured_image",
    "main_image",
    "product_image",
    "image_url",
    "image_src"

  ];

  for (
    let i = 0;
    i < possibleFields.length;
    i++
  ) {

    const value =
      getXmlValue(
        item,
        possibleFields[i]
      );

    if (
      value &&
      isValidImageUrl(value)
    ) {

      return cleanImageUrl(value);

    }

  }

  // =====================================================
  // Search XML children for image-like fields
  // =====================================================

  const children =
    item.getChildren();

  for (
    let i = 0;
    i < children.length;
    i++
  ) {

    const child =
      children[i];

    const name =
      normalizeName(
        child.getName()
      );

    const text =
      clean(
        child.getText()
      );

    if (
      !text
    ) {
      continue;
    }

    if (
      name.indexOf("image") !== -1 &&
      isValidImageUrl(text)
    ) {

      return cleanImageUrl(text);

    }

  }

  return "";

}


// =====================================================
// VALIDATE IMAGE URL
// =====================================================

function isValidImageUrl(url) {

  if (!url) {
    return false;
  }

  const value =
    String(url)
      .trim();

  if (!value) {
    return false;
  }

  return (
    /^https?:\/\//i.test(value)
  );

}


// =====================================================
// CLEAN IMAGE URL
// =====================================================

function cleanImageUrl(url) {

  if (!url) {
    return "";
  }

  let value =
    String(url)
      .trim();

  // Remove accidental whitespace
  value =
    value.replace(
      /\s+/g,
      ""
    );

  return value;

}


// =====================================================
// CREATE BASE HANDLE FROM URL
// =====================================================

function createBaseHandleFromUrl(url) {

  if (!url) {
    return "";
  }

  let cleanUrl =
    String(url)
      .trim();

  // -----------------------------------------------------
  // Remove hash
  // -----------------------------------------------------

  const hashPosition =
    cleanUrl.indexOf("#");

  if (hashPosition !== -1) {

    cleanUrl =
      cleanUrl.substring(
        0,
        hashPosition
      );

  }

  // -----------------------------------------------------
  // Remove query
  // -----------------------------------------------------

  let path =
    cleanUrl;

  const questionMark =
    cleanUrl.indexOf("?");

  if (questionMark !== -1) {

    path =
      cleanUrl.substring(
        0,
        questionMark
      );

  }

  // -----------------------------------------------------
  // Remove trailing slash
  // -----------------------------------------------------

  path =
    path.replace(
      /\/+$/,
      ""
    );

  // -----------------------------------------------------
  // Final URL segment
  // -----------------------------------------------------

  const parts =
    path.split("/");

  let slug =
    parts[
      parts.length - 1
    ];

  try {

    slug =
      decodeURIComponent(slug);

  } catch (e) {}

  // -----------------------------------------------------
  // Remove packaging
  // -----------------------------------------------------

  slug =
    removeExistingPackagingSuffix(
      slug
    );

  // -----------------------------------------------------
  // Shopify handle
  // -----------------------------------------------------

  return createShopifyHandle(
    slug
  );

}


// =====================================================
// REMOVE PACKAGING FROM SLUG
// =====================================================

function removeExistingPackagingSuffix(slug) {

  if (!slug) {
    return "";
  }

  let result =
    String(slug)
      .trim()
      .toLowerCase();

  // =====================================================
  // Multipack decimal
  //
  // 3-x-1-5-kg
  // =====================================================

  result =
    result.replace(
      /-\d+-x-\d+-\d+-(?:kg|kgr|g|gr|ml|l)$/i,
      ""
    );

  // =====================================================
  // Multipack whole
  //
  // 6-x-5-kg
  // =====================================================

  result =
    result.replace(
      /-\d+-x-\d+-(?:kg|kgr|g|gr|ml|l)$/i,
      ""
    );

  // =====================================================
  // 2-1-5-kgr
  // =====================================================

  result =
    result.replace(
      /-\d+-\d+-\d+-(?:kg|kgr|g|gr|ml|l)$/i,
      ""
    );

  // =====================================================
  // 2x1-5-kg
  // =====================================================

  result =
    result.replace(
      /-\d+x\d+-\d+-(?:kg|kgr|g|gr|ml|l)$/i,
      ""
    );

  // =====================================================
  // 6x185gr
  // =====================================================

  result =
    result.replace(
      /-\d+x\d+(?:kg|kgr|g|gr|ml|l)$/i,
      ""
    );

  // =====================================================
  // 1-5-kg
  // =====================================================

  result =
    result.replace(
      /-\d+-\d+-(?:kg|kgr|g|gr|ml|l)$/i,
      ""
    );

  // =====================================================
  // 10-kg
  // =====================================================

  result =
    result.replace(
      /-\d+-(?:kg|kgr|g|gr|ml|l)$/i,
      ""
    );

  // =====================================================
  // 1.5kg
  // =====================================================

  result =
    result.replace(
      /-\d+\.\d+(?:kg|kgr|g|gr|ml|l)$/i,
      ""
    );

  // =====================================================
  // 10kg
  // =====================================================

  result =
    result.replace(
      /-\d+(?:kg|kgr|g|gr|ml|l)$/i,
      ""
    );

  return result;

}


// =====================================================
// SHOPIFY HANDLE
// =====================================================

function createShopifyHandle(text) {

  if (!text) {
    return "";
  }

  let value =
    String(text)
      .trim()
      .toLowerCase();

  // =====================================================
  // Remove accents
  // =====================================================

  value =
    value.normalize("NFD")
      .replace(
        /[\u0300-\u036f]/g,
        ""
      );

  // =====================================================
  // Greek transliteration
  // =====================================================

  const greekMap = {

    "α": "a",
    "β": "v",
    "γ": "g",
    "δ": "d",
    "ε": "e",
    "ζ": "z",
    "η": "i",
    "θ": "th",
    "ι": "i",
    "κ": "k",
    "λ": "l",
    "μ": "m",
    "ν": "n",
    "ξ": "x",
    "ο": "o",
    "π": "p",
    "ρ": "r",
    "σ": "s",
    "ς": "s",
    "τ": "t",
    "υ": "y",
    "φ": "f",
    "χ": "ch",
    "ψ": "ps",
    "ω": "o"

  };

  value =
    value.replace(
      /[α-ω]/g,
      function(letter) {

        return (
          greekMap[letter] ||
          letter
        );

      }
    );

  // =====================================================
  // Shopify characters
  // =====================================================

  value =
    value.replace(
      /[^a-z0-9\s-]/g,
      ""
    );

  // =====================================================
  // Spaces → hyphens
  // =====================================================

  value =
    value.replace(
      /\s+/g,
      "-"
    );

  // =====================================================
  // Duplicate hyphens
  // =====================================================

  value =
    value.replace(
      /-+/g,
      "-"
    );

  // =====================================================
  // Trim hyphens
  // =====================================================

  value =
    value.replace(
      /^-+|-+$/g,
      ""
    );

  // =====================================================
  // Shopify max length
  // =====================================================

  if (value.length > 255) {

    value =
      value.substring(
        0,
        255
      );

    value =
      value.replace(
        /-+$/g,
        ""
      );

  }

  return value;

}


// =====================================================
// FIND XML ITEMS
// =====================================================

function collectItems(element, items) {

  const name =
    String(
      element.getName()
    ).toLowerCase();

  if (
    name === "item"
  ) {

    items.push(element);

    return;

  }

  const children =
    element.getChildren();

  children.forEach(function(child) {

    collectItems(
      child,
      items
    );

  });

}


// =====================================================
// GET XML VALUE
// =====================================================

function getXmlValue(item, fieldName) {

  if (!item) {
    return "";
  }

  const children =
    item.getChildren();

  const wanted =
    normalizeName(
      fieldName
    );

  for (
    let i = 0;
    i < children.length;
    i++
  ) {

    const child =
      children[i];

    const childName =
      normalizeName(
        child.getName()
      );

    if (
      childName === wanted
    ) {

      return clean(
        child.getText()
      );

    }

  }

  return "";

}


// =====================================================
// NORMALIZE XML FIELD NAME
// =====================================================

function normalizeName(name) {

  return String(name)
    .toLowerCase()
    .replace(
      /^g:/,
      ""
    )
    .replace(
      /[_\-\s]/g,
      ""
    );

}


// =====================================================
// PRICE PARSER
// =====================================================

function parsePrice(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {

    return "";

  }

  let text =
    String(value)
      .trim();

  // =====================================================
  // Remove currency
  // =====================================================

  text =
    text.replace(
      /[^\d,.-]/g,
      ""
    );

  if (!text) {
    return "";
  }

  // =====================================================
  // 1.299,95
  // =====================================================

  if (
    text.indexOf(".") !== -1 &&
    text.indexOf(",") !== -1
  ) {

    if (
      text.lastIndexOf(",") >
      text.lastIndexOf(".")
    ) {

      text =
        text
          .replace(
            /\./g,
            ""
          )
          .replace(
            ",",
            "."
          );

    } else {

      text =
        text.replace(
          /,/g,
          ""
        );

    }

  }

  // =====================================================
  // 39,86
  // =====================================================

  else if (
    text.indexOf(",") !== -1
  ) {

    text =
      text.replace(
        ",",
        "."
      );

  }

  const number =
    Number(text);

  if (
    isNaN(number)
  ) {

    return "";

  }

  return number;

}


// =====================================================
// CLEAN
// =====================================================

function clean(value) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }

  return String(value)
    .trim();

}


// =====================================================
// REMOVE HTML
// =====================================================

function stripHTML(value) {

  if (!value) {
    return "";
  }

  return String(value)
    .replace(
      /<[^>]*>/g,
      " "
    )
    .replace(
      /\s+/g,
      " "
    )
    .trim();

}