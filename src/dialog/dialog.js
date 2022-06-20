const roamingSelect = document.getElementById("roaming-select");
const roamingSet = document.getElementById("roaming-set");
const roamingCurrent = document.getElementById("roaming-current");

Office.onReady(() => {
  const initRoamingVal = Office.context.roamingSettings.get("roaming-key");
  roamingCurrent.innerText = initRoamingVal == (null || undefined) ? "Not set" : initRoamingVal;

  roamingSet.addEventListener("click", () => {
    Office.context.roamingSettings.set("roaming-key", roamingSelect.value);
    Office.context.roamingSettings.saveAsync(() => {
      console.log("Roaming settings saved");
      roamingCurrent.innerText = Office.context.roamingSettings.get("roaming-key");
    });
  });
});