import scrapy.cmdline

if __name__ == "__main__":
    print("HELLO, SPIDER!")
    # scrapy.cmdline._run_command("scrapy crawl place",
    #                             args={"country": "Sri Lanka"},
    #                             opts="")
    scrapy.cmdline.execute(
        "scrapy crawl place -a country=Sri-Lanka -a interests=Establishment".
        split())
    #scrapy.cmdline.execute("scrapy crawl place a country=Sri Lanka".split())
