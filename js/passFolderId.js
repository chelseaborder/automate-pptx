let dropdown = document.getElementById('select-box')

for (let i = 0; i < folder.length; i++) {
  //Only display sub-folders
  if (folder[i].childIds.length > 0) {
    console.log()
  } else {
    option = document.createElement('option');
    option.text = folder[i].name;
    option.value = folder[i].id
    dropdown.add(option);
  }
};

function getValue() {
  let selectedFolderValue = dropdown.options[dropdown.selectedIndex].value
  console.log("selected value: " + selectedFolderValue)

}
