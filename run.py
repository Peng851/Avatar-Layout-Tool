from image_arranger import ImageArranger

if __name__ == "__main__":
    try:
        app = ImageArranger()
        app.run()
    except Exception as e:
        print(f"程序运行出错: {e}")
        input("按回车键退出...") 