/**
 * [실행 명령어] node --max-old-space-size=4096 index.js
 * [url수정, 파일이름수정, sheet 이름 수정]
 * [issue] 'out of memory' 문제로 약 450 count 마다 새로고침해줘야 하는 문제가 있음.
 * https://www.tutorialspoint.com/how-to-solve-process-out-of-memory-exception-in-node-js
 */
const { Builder, By, Key, until, Capabilities } = require("selenium-webdriver")
const chrome = require("selenium-webdriver/chrome")
const XLSX = require("xlsx")
const fs = require("fs")
const capabilities = Capabilities.chrome().setAcceptInsecureCerts(true)
let array = []
const run = async () => {
  let driver = await new Builder()
    .forBrowser("chrome")
    .withCapabilities(capabilities)
    .setChromeOptions(
      new chrome.Options().addArguments(
        "--headless",
        "--ignore-certificate-error",
        "--ignore-ssl-errors",
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
      //test
      // if (totalCount >= 10) break
      await driver.get(
        `https://www.karhanbang.com/office/office_list.asp?topM=09&flag=G&page=${nowPage++}&search=&sel_sido=2&sel_gugun=28&sel_dong=`
      )

      await driver.executeScript("window.scroll(0," + 500 + ");")
      // let pageSize = //*[@id="contents"]/div[1]/div[3]/div/div/a[1] //*[@id="contents"]/div[1]/div[3]/div/div/a[2] //*[@id="contents"]/div[1]/div[3]/div/div/a[3]

      let rowCount = (
        await driver.findElements(
          By.xpath('//*[@id="contents"]/div[1]/div[3]/div/table/tbody/tr')
        )
      ).length
      console.log("page: ", nowPage - 1)
      console.log("row count: ", rowCount)

      if (rowCount == 0) break
      for (let i = 1; i <= rowCount; i++) {
        let region1 = await driver.wait(
          until.elementLocated(
            By.xpath(
              `//*[@id="contents"]/div[1]/div[3]/div/table/tbody/tr[${i}]/td[1]`
            )
          )
        )
        const region = (await region1.getText()).split(" ")[0]

        await driver
          .wait(
            until.elementLocated(
              By.xpath(
                `//*[@id="contents"]/div[1]/div[3]/div/table/tbody/tr[${i}]/td[2]/p/a`
              )
            )
          )
          .click()

        let company1 = await driver.wait(
          until.elementLocated(
            By.xpath('//*[@id="realtorsDetail"]/div[1]/div[3]/dl/dt')
          )
        )
        const company = (await company1.getText()).split("\n")[0]

        let name1 = await driver.wait(
          until.elementLocated(
            By.xpath(
              '//*[@id="realtorsDetail"]/div[1]/div[3]/dl/dd/ul/li[2]/div/ul/li[1]/em'
            )
          )
        )
        const name = await name1.getText()

        let address1 = await driver.wait(
          until.elementLocated(
            By.xpath(
              '//*[@id="realtorsDetail"]/div[1]/div[3]/dl/dd/ul/li[2]/div/ul/li[2]/em'
            )
          )
        )
        const address = await address1.getText()

        let tel1 = await driver.wait(
          until.elementLocated(
            By.xpath(
              '//*[@id="realtorsDetail"]/div[1]/div[3]/dl/dd/ul/li[2]/div/ul/li[5]/em/a[1]/font'
            )
          )
        )
        const tel = await tel1.getText()

        let phone1 = await driver.wait(
          until.elementLocated(
            By.xpath(
              '//*[@id="realtorsDetail"]/div[1]/div[3]/dl/dd/ul/li[2]/div/ul/li[5]/em/a[2]/font'
            )
          )
        )
        const phone = await phone1.getText()

        //console.log("상호: ", company)
        // console.log("name: ", name);
        // console.log("address: ", address);
        // console.log("전화번호: ", tel);
        // console.log("휴대폰번호: ", phone);
        totalCount++
        console.log("totalCount: ", totalCount)
        array.push({
          지역: region,
          상호: company,
          이름: name,
          주소: address,
          전화번호: tel,
          핸드폰번호: phone,
        })
        await driver.navigate().back()
      }
    }
    // console.log(array)
    // console.log("전체: ", totalCount);

    const ws = XLSX.utils.json_to_sheet(array)

    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, "김포시")

    const excelFileName = "김포시.xlsx"
    XLSX.writeFile(wb, excelFileName)
  } catch (e) {
    console.error(e)
  } finally {
    driver.quit()
  }
}

run()
