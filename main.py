import SetupCoupon as SC
import DataSource as DS


def main():
    # 打开浏览器
    url = "https://jsls.jinritemai.com/mfa/retail-marketing/tools/product-coupons/list"
    driver = SC.open_browser(url)
    # 读取设券模板表
    source = r'D:\移动终端广分互联网\小时达\其他\设券模板表.xlsx'
    df = DS.read_excel(source)
    # 循环遍历每一行
    try:
        for i in range(len(df)):
            coupon_list = df.iloc[i]
            print(coupon_list)

            # 设置满减券
            SC.setup_coupon(driver, coupon_list)
            driver.get(url)
    except Exception as e:
        print(e)
    finally:
        SC.close_browser(driver)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

