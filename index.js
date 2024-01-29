const { Builder, By, Key, until, Capabilities } = require("selenium-webdriver")
const chrome = require("selenium-webdriver/chrome")
const XLSX = require("xlsx")
const fs = require("fs")
const capabilities = Capabilities.chrome()
let array = []
const run = async () => {
  let driver = await new Builder()
    .forBrowser("chrome")
    .withCapabilities(capabilities)
    .setChromeOptions(
      new chrome.Options().addArguments(
        "--disable-gpu",
        "window-size=1920x1080",
        "lang=ko_KR"
      )
    )
    .build()

  try {
    let nowPage = 1
    let totalCount = 0
    while (true) {
      await driver.get(
        `https://www.karhanbang.com/office/office_list.asp?topM=09&flag=G&page=${nowPage++}&search=&sel_sido=10&sel_gugun=67&sel_dong=`
      )
      await driver.executeScript("window.scroll(0," + 500 + ");")
      // let pageSize = //*[@id="contents"]/div[1]/div[3]/div/div/a[1] //*[@id="contents"]/div[1]/div[3]/div/div/a[2] //*[@id="contents"]/div[1]/div[3]/div/div/a[3]
      let rowCount = (
        await driver.findElements(
          By.xpath('//*[@id="contents"]/div[1]/div[3]/div/table/tbody/tr')
        )
      ).length
      if (rowCount == 0) break
      for (let i = 1; i <= rowCount; i++) {
        await driver
          .wait(
            until.elementLocated(
              By.xpath(
                `//*[@id="contents"]/div[1]/div[3]/div/table/tbody/tr[${i}]/td[2]`
              )
            )
          )
          .click()

        let name1 = await driver.wait(
          until.elementLocated(
            By.xpath(
              '//*[@id="realtorsDetail"]/div[1]/div[3]/dl/dd/ul/li[2]/div/ul/li[1]/em'
            )
          )
        )
        let address1 = await driver.wait(
          until.elementLocated(
            By.xpath(
              '//*[@id="realtorsDetail"]/div[1]/div[3]/dl/dd/ul/li[2]/div/ul/li[2]/em'
            )
          )
        )
        let tel1 = await driver.wait(
          until.elementLocated(
            By.xpath(
              '//*[@id="realtorsDetail"]/div[1]/div[3]/dl/dd/ul/li[2]/div/ul/li[5]/em/a[1]/font'
            )
          )
        )
        let phone1 = await driver.wait(
          until.elementLocated(
            By.xpath(
              '//*[@id="realtorsDetail"]/div[1]/div[3]/dl/dd/ul/li[2]/div/ul/li[5]/em/a[2]/font'
            )
          )
        )

        const name = await name1.getText()
        const address = await address1.getText()
        const tel = await tel1.getText()
        const phone = await phone1.getText()
        // console.log("name: ", name);
        // console.log("address: ", address);
        // console.log("전화번호: ", tel);
        // console.log("휴대폰번호: ", phone);
        totalCount++
        array.push({ name, address, tel, phone })
        await driver.navigate().back()
      }
    }
    // console.log(array)
    // console.log("전체: ", totalCount);

    const ws = XLSX.utils.json_to_sheet(array)

    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1")

    const excelFileName = "data.xlsx"
    XLSX.writeFile(wb, excelFileName)
  } catch (e) {
    console.error(e)
  } finally {
    driver.quit()
  }
}

run()
