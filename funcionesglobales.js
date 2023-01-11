function colorCountryTab(tab,country) {
    /*
    Esta función necesita como inputs la tab y el país, y directamente colorea la tab
    */
    const countryColors = {
    Argentina: "#00FF00",//Argentina: Green
    Australia: "#0000FF",//Australia: Blue
    Austria: "#FFFF00",//Austria: Yellow
    Belgium: "#FF00FF",//Belgium: Fuchsia
    Brazil: "#00FFFF",//Brazil: Aqua
    Canada: "#FF0000",//Canada: Red
    Chile: "#800000",//Chile: Maroon
    "Costa Rica": "#008000",//Costa Rica: Olive
    Denmark: "#000080",//Denmark: Navy
    France: "#808000",//France: Olive
    Germany: "#800080",//Germany: Purple
    Ireland: "#008080",//Ireland: Teal
    Italy: "#808080",//Italy: Gray
    Japan: "#C0C0C0",//Japan: Silver
    "Korea (Republic of)": "#FFC0CB",//Korea (Republic of): Pink
    Mexico: "#FFA07A",//Mexico: Light Salmon
    Morocco: "#FFD700",//Morocco: Gold
    Netherlands: "#DAA520",//Netherlands: Goldenrod
    "New Zealand": "#ADD8E6",//New Zealand: Light Blue
    Portugal: "#98FB98",//Portugal: Pale Green
    "Russian Federation": "#F0E68C",//Russian Federation: Khaki
    Singapore: "#E6E6FA",//Singapore: Lavender
    "South Africa": "#D3D3D3",//South Africa: Light Gray
    Spain: "#A52A2A",//Spain: Brown
    Sweden: "#FF69B4",//Sweden: Hot Pink
    Switzerland: "#FF1493",//Switzerland: Deep Pink
    "United Arab Emirates": "#DB7093",//United Arab Emirates: Pale Violet Red
    "United Kingdom": "#B0C4DE",//United Kingdom: Light Steel Blue
    "United States": "#87CEEB"//United States: Sky Blue
    };
    tab.setTabColor(countryColors[country] || "#000000");
  }

  function checkInputs(){
    
  }