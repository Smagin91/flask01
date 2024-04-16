import requests
import json
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from prophet import Prophet
import random
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from flask import Flask, render_template, request, redirect, send_from_directory, send_file, flash
from werkzeug.utils import secure_filename
import os
import numpy as np
import threading
import time
from openpyxl import Workbook


app = Flask(__name__)
app.secret_key = "super_secret_key"
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}
parsing_status = False
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('Не указан путь файла')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('Не выбран файл')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            try:
                df = pd.read_excel(file_path)
                if 'ds' not in df.columns or 'y' not in df.columns:
                    flash('Отсутствует столбец дат "ds" или показателей "y" с заголовках.')
                    return redirect(request.url)
                if len(df) < 36:
                    flash('Недостаточно точек для прогноза. Требуется минимум 36 наблюдений.')
                    return redirect(request.url)
                # Проверяем даты

                # Обучаем модель
                prophet_model = Prophet()
                prophet_model.fit(df)
                future = prophet_model.make_future_dataframe(periods=6, freq='M')
                forecast = prophet_model.predict(future)

                # Plotting
                plt.figure(figsize=(11, 5))

                # Определим количество точек для отображения
                num_points_to_display = 36
                # Возьмем последние num_points_to_display точек из исходных данных
                df_to_display = df.iloc[-num_points_to_display:]
                # Построим график для этих точек
                plt.plot(df_to_display['ds'], df_to_display['y'], label='Исторические данные')

                # Добавим прогноз на 6 точек вперед
                future_dates = forecast['ds'].iloc[-6:]
                plt.plot(future_dates, forecast['yhat'].iloc[-6:], linestyle='--', label='Прогноз')

                plt.xlabel('Дата')
                plt.ylabel('ИПЦ, mom')
                plt.title('Прогноз на 6 месяцев')
                plt.legend()
                recent_forecast_plot_path = 'static/recent_forecast_plot.png'
                plt.savefig(recent_forecast_plot_path)
                plt.close()

                # Строим график с доверительным интервалом
                plt.figure(figsize=(10, 5))
                prophet_model.plot(forecast)
                plt.xlabel('Дата')
                plt.ylabel('ИПЦ, mom')
                plt.title('Весь период и доверительный интервал')
                plt.legend(['Факт', 'Прогноз', 'Доверительный интервал'])
                plt.tight_layout(rect=[0, 0.03, 1, 0.95])
                complete_forecast_plot_path = 'static/complete_forecast_plot.png'
                plt.savefig(complete_forecast_plot_path)
                plt.close()

                # СОхраняем прогноз в эксель
                forecast.to_excel('static/forecast.xlsx', index=False)

                return render_template('result.html', recent_plot_url=recent_forecast_plot_path,
                                       complete_plot_url=complete_forecast_plot_path)

            except Exception as e:
                flash(f'Error processing file: {str(e)}')
                return redirect(request.url)

        else:
            flash('Неверный формат файла. Поддерживаются только .xlsx файлы excel.')
            return redirect(request.url)
    return render_template('index.html')


@app.route('/auto', methods=['GET', 'POST'])
def auto():
    if request.method == 'POST':
        df = pd.DataFrame(columns=['Дата', 'Код товара или услуги', 'Название товара или услуги', 'Месячный прирост'])

        # Задаем начальную и конечную даты для цикла
        start_date = datetime(2002, 1, 1)
        # Округляем текущую дату до начала месяца
        end_date = datetime.now().replace(day=1)
        # end_date = datetime(2012, 1, 1)
        # Вычитаем один месяц
        end_date -= relativedelta(months=1)

        # Итерируемся по датам
        current_date = start_date
        while current_date <= end_date:
            try:
                # Формируем URL с текущей датой
                url = f'https://showdata.gks.ru/x/report/277326/view/compound/?&filter_1_0={current_date.strftime("%Y-%m-%d")}+00%3A00%3A00%7C-56&filter_2_0=120472&filter_3_0=13035&filter_4_0=109481%2C109485%2C109513%2C109420%2C109543%2C347484%2C109515%2C109393%2C109516%2C109363%2C109394%2C109421%2C109422%2C109545%2C109423%2C109546%2C109424%2C109547%2C109425%2C109548%2C109517%2C109426%2C109549%2C109427%2C109550%2C109428%2C109551%2C318638%2C347550%2C109395%2C109429%2C109552%2C109430%2C339613%2C339769%2C339770%2C339771%2C347225%2C347228%2C109518%2C109553%2C109396%2C109431%2C109519%2C109397%2C109520%2C217799%2C109486%2C109398%2C109554%2C109432%2C109555%2C109433%2C347487%2C109521%2C109556%2C109434%2C109557%2C109435%2C109558%2C109436%2C109399%2C109522%2C109364%2C109400%2C109559%2C109437%2C109560%2C109523%2C109438%2C109561%2C109439%2C109562%2C109440%2C109563%2C109441%2C109565%2C217792%2C109443%2C109566%2C109444%2C109445%2C109568%2C109446%2C109569%2C109447%2C109570%2C109448%2C109571%2C109449%2C109572%2C109573%2C109451%2C109574%2C109452%2C288012%2C109575%2C109453%2C109359%2C109487%2C109401%2C109454%2C109577%2C109455%2C109578%2C109456%2C339773%2C347231%2C109524%2C109402%2C109525%2C109403%2C109526%2C109404%2C109527%2C109405%2C109365%2C109528%2C109579%2C109457%2C109488%2C109406%2C109581%2C109459%2C109582%2C109460%2C109583%2C109461%2C109584%2C339606%2C109462%2C109529%2C109585%2C109463%2C109586%2C109464%2C109588%2C109466%2C109589%2C109467%2C109366%2C109590%2C109468%2C109469%2C109489%2C109470%2C109593%2C109471%2C322817%2C109594%2C109472%2C109595%2C109596%2C109474%2C109597%2C318637%2C109475%2C109598%2C109476%2C109599%2C109477%2C109600%2C109478%2C109601%2C109479%2C109602%2C318636%2C109480%2C109726%2C109604%2C109605%2C109728%2C109606%2C109729%2C109607%2C110321%2C347014%2C109730%2C109608%2C109731%2C347475%2C347485%2C109482%2C109367%2C109407%2C109609%2C109610%2C109733%2C109530%2C217794%2C339775%2C109611%2C109612%2C109735%2C109613%2C109736%2C109614%2C339614%2C109737%2C109738%2C109739%2C109617%2C109740%2C109618%2C109741%2C109619%2C109742%2C109620%2C109621%2C109744%2C109622%2C110322%2C109745%2C109624%2C109747%2C109625%2C109626%2C109749%2C109627%2C109750%2C109628%2C109751%2C109360%2C109490%2C109408%2C109629%2C109752%2C109531%2C109532%2C109410%2C109533%2C109630%2C109753%2C109631%2C109754%2C109411%2C109368%2C109632%2C109534%2C109755%2C109633%2C109756%2C109757%2C109635%2C109758%2C109412%2C109759%2C109637%2C109760%2C109638%2C109761%2C109639%2C109762%2C109640%2C109763%2C109641%2C109535%2C109764%2C109642%2C109765%2C109643%2C347476%2C109644%2C109645%2C217795%2C109646%2C109769%2C109647%2C109770%2C110424%2C109771%2C109649%2C109772%2C217793%2C109773%2C109651%2C109774%2C109652%2C109775%2C109653%2C109654%2C109777%2C109655%2C109778%2C110325%2C288015%2C109657%2C109780%2C109658%2C109781%2C109659%2C109783%2C109661%2C109784%2C109663%2C109786%2C109664%2C109666%2C109789%2C109667%2C109790%2C109668%2C109791%2C109491%2C109413%2C109669%2C109672%2C109795%2C109674%2C109797%2C109799%2C109677%2C109801%2C109369%2C109536%2C109679%2C109680%2C109803%2C109681%2C109804%2C347223%2C109682%2C109805%2C109683%2C109806%2C109684%2C109685%2C109808%2C291942%2C109687%2C109810%2C109688%2C109811%2C109689%2C109812%2C109691%2C347478%2C109814%2C109692%2C109693%2C109816%2C109694%2C109817%2C109695%2C288010%2C109370%2C109818%2C109696%2C109697%2C109820%2C109821%2C109823%2C109701%2C109824%2C109702%2C109825%2C109704%2C109827%2C339607%2C339615%2C109828%2C109829%2C109831%2C109709%2C109832%2C109710%2C109833%2C109834%2C109712%2C109835%2C109713%2C318625%2C109714%2C109837%2C109715%2C109838%2C339611%2C109493%2C109716%2C109717%2C109840%2C109371%2C109483%2C109494%2C109414%2C109841%2C109537%2C109415%2C109372%2C109842%2C109720%2C318631%2C109495%2C109721%2C109373%2C109844%2C109722%2C109723%2C110373%2C109847%2C110374%2C347230%2C109848%2C110375%2C109849%2C110376%2C109850%2C110377%2C109851%2C110327%2C110435%2C318628%2C288014%2C318649%2C318635%2C109496%2C110378%2C109852%2C110379%2C109853%2C347232%2C110380%2C109854%2C288009%2C110381%2C110436%2C109855%2C110382%2C288016%2C110384%2C109858%2C339608%2C109374%2C110385%2C109859%2C109497%2C110386%2C109860%2C110387%2C109861%2C110388%2C339616%2C109375%2C109862%2C110389%2C109498%2C109863%2C110390%2C109864%2C110391%2C110393%2C217798%2C109868%2C110395%2C109869%2C110396%2C109870%2C110397%2C109871%2C110398%2C109872%2C217796%2C110399%2C109873%2C110400%2C318639%2C318646%2C109376%2C109874%2C110401%2C109875%2C109361%2C109538%2C110402%2C109876%2C110403%2C109877%2C110404%2C109416%2C347488%2C347479%2C347489%2C347480%2C109539%2C109417%2C339774%2C339768%2C109878%2C109879%2C110406%2C109880%2C110407%2C347490%2C109881%2C109882%2C110409%2C109883%2C109966%2C109884%2C109886%2C109887%2C109970%2C109888%2C109971%2C109889%2C109972%2C109890%2C109973%2C109891%2C109974%2C109892%2C109893%2C109977%2C109895%2C109978%2C109896%2C109979%2C109897%2C109980%2C109898%2C109981%2C288013%2C347481%2C110330%2C110331%2C110332%2C109899%2C109982%2C109900%2C109983%2C109901%2C318643%2C318621%2C318641%2C318647%2C347226%2C109984%2C109902%2C109903%2C109986%2C109904%2C109987%2C109906%2C109989%2C109907%2C109990%2C347491%2C339609%2C109908%2C109991%2C109909%2C109992%2C109484%2C109377%2C109993%2C109913%2C109996%2C109914%2C109997%2C109418%2C109915%2C109998%2C109917%2C109918%2C110001%2C109919%2C110002%2C109920%2C110003%2C110333%2C288011%2C109500%2C109921%2C110004%2C110005%2C110335%2C318648%2C347235%2C109378%2C109923%2C109924%2C109925%2C110008%2C109926%2C110010%2C109928%2C110011%2C109929%2C347486%2C110012%2C109931%2C347227%2C110014%2C110015%2C110016%2C109934%2C110017%2C109935%2C110018%2C109936%2C110019%2C109937%2C110020%2C109938%2C110021%2C109939%2C110022%2C110336%2C110437%2C109940%2C109502%2C110023%2C110024%2C109380%2C109503%2C110025%2C110026%2C109944%2C109945%2C110028%2C109946%2C217805%2C347233%2C110337%2C110029%2C109947%2C110030%2C109948%2C110031%2C110338%2C217807%2C217787%2C109949%2C110034%2C109952%2C109953%2C110036%2C318661%2C110037%2C110038%2C109956%2C110039%2C109957%2C110040%2C109958%2C109959%2C318665%2C110042%2C109960%2C110043%2C109961%2C110044%2C109962%2C110045%2C109963%2C110046%2C318657%2C318653%2C318658%2C318652%2C322818%2C318664%2C318666%2C109964%2C339633%2C318668%2C318663%2C318662%2C339620%2C339634%2C339621%2C339635%2C339622%2C339636%2C339623%2C339619%2C339637%2C339624%2C110047%2C109965%2C110048%2C110370%2C110049%2C110371%2C339638%2C110372%2C110174%2C110052%2C110175%2C110053%2C110176%2C339625%2C110054%2C110177%2C110055%2C110178%2C110056%2C110179%2C110057%2C110180%2C110058%2C339639%2C110181%2C110059%2C110182%2C110060%2C110183%2C110061%2C110184%2C110062%2C110185%2C110063%2C110186%2C110064%2C110187%2C110065%2C110188%2C110066%2C318667%2C339626%2C339612%2C109362%2C109381%2C347229%2C347221%2C110069%2C110192%2C109419%2C110070%2C110193%2C217803%2C318623%2C318640%2C318618%2C318626%2C339640%2C339627%2C109542%2C339641%2C339628%2C339642%2C339629%2C339643%2C339630%2C339644%2C339631%2C339645%2C339632%2C339772%2C347482%2C347492%2C347483%2C109504%2C110071%2C110194%2C110072%2C109382%2C110195%2C110073%2C110196%2C110074%2C109505%2C110197%2C110075%2C110198%2C110076%2C110199%2C110077%2C109383%2C109506%2C109384%2C109507%2C109385%2C109508%2C109386%2C110200%2C110078%2C110201%2C110079%2C110202%2C110080%2C110081%2C110204%2C110082%2C110083%2C110208%2C110086%2C110209%2C110087%2C110210%2C347477%2C110088%2C110211%2C110089%2C110212%2C110090%2C110091%2C110214%2C110092%2C110093%2C110216%2C110094%2C110217%2C110095%2C110218%2C110096%2C110219%2C110098%2C110221%2C110438%2C217797%2C318624%2C339617%2C110339%2C110340%2C110101%2C110224%2C110102%2C110225%2C110103%2C110226%2C110104%2C110227%2C110105%2C110228%2C110341%2C110106%2C217806%2C318617%2C110230%2C110108%2C110231%2C110110%2C110233%2C110111%2C110234%2C110112%2C110235%2C110113%2C347234%2C347548%2C347549%2C110236%2C110114%2C110237%2C217801%2C110115%2C110238%2C110116%2C110239%2C110117%2C110240%2C110241%2C110119%2C110242%2C110120%2C110243%2C110121%2C110244%2C110122%2C110245%2C110123%2C110246%2C217788%2C110124%2C110125%2C110343%2C318616%2C110248%2C110344%2C110345%2C343206%2C110126%2C352389%2C110250%2C110128%2C110251%2C110129%2C110252%2C110130%2C110253%2C110131%2C110254%2C110132%2C110133%2C110439%2C110135%2C110258%2C110136%2C110259%2C110137%2C110260%2C110138%2C110261%2C110139%2C284143%2C284144%2C110262%2C110140%2C110263%2C110141%2C110264%2C110142%2C110346%2C110265%2C110347%2C318633%2C318634%2C318614%2C318619%2C318622%2C110143%2C110266%2C110144%2C110267%2C110145%2C110268%2C110146%2C352390%2C110269%2C110147%2C110270%2C110149%2C110440%2C110272%2C110150%2C110151%2C110152%2C110153%2C347224%2C110277%2C318630%2C347222%2C110155%2C110278%2C110156%2C110279%2C110157%2C110280%2C110158%2C110281%2C110416%2C110441%2C110159%2C318615%2C339610%2C343217%2C343228%2C343210%2C343215%2C343222%2C343226%2C343219%2C110282%2C110160%2C110162%2C110285%2C110163%2C110164%2C110287%2C110165%2C110288%2C217802%2C110290%2C110168%2C110169%2C110292%2C110170%2C110293%2C110171%2C110294%2C110295%2C318629%2C318632%2C318620%2C110354%2C110296%2C110355%2C110297%2C217800%2C110356%2C110298%2C110357%2C110299%2C110358%2C110300%2C318642%2C110359%2C110301%2C110360%2C288017%2C110302%2C110361%2C110303%2C339618%2C217804%2C110362%2C110304%2C110363%2C110305%2C110364%2C110306%2C110365%2C110307%2C110366%2C110308%2C110367%2C110309%2C110368%2C110310%2C110369%2C110352&rp_submit=t&_=1713022138088'

                # Отправляем запрос и получаем данные
                response = requests.get(url, verify=False).json()

                # Инициализируем списки для хранения данных
                db_names = []
                db_values = []
                db_codes = []

                # Извлекаем данные из ответа
                for item in response['data']['reportData']['data']:
                    db_values.append(item[0]['db_value'])

                for item in response['headers']['reportHeaders']['row_header']:
                    for child in item['children']:
                        for sub_child in child['children']:
                            db_names.append(sub_child['display_title'])

                for item in response['headers']['reportHeaders']['row_header']:
                    for child in item['children']:
                        for sub_child in child['children']:
                            db_codes.append(sub_child['extra_row_attrs']['code'])

                # Создаем временный датафрейм с данными за текущую дату
                temp_df = pd.DataFrame({
                    'Дата': [current_date.strftime('%Y-%m-%d')] * len(db_codes),
                    'Код товара или услуги': db_codes,
                    'Название товара или услуги': db_names,
                    'Месячный прирост': db_values,
                })

                # Добавляем временный датафрейм к основному
                df = pd.concat([df, temp_df], ignore_index=True)

                # Уменьшаем текущую дату на один месяц
                current_date += relativedelta(months=1)
            except Exception as e:
                print(f"Произошла ошибка на дате {current_date.strftime('%Y-%m-%d')}: {e}")
                print("Повторная попытка через 10 секунд...")
        # Повернем таблицу, используя коды в качестве заголовков столбцов
        pivot_df = df.pivot_table(index='Дата', columns='Код товара или услуги', values='Месячный прирост')


        # Сохраняем данные в файл xlsx
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'data.xlsx')
        pivot_df.to_excel(file_path, index=True)

        return render_template('auto.html', parsed=True, file_path=file_path)

    return render_template('auto.html', parsed=False)

@app.route('/download_data')
def download_data():
    file_path = request.args.get('file_path')
    return send_from_directory(app.config['UPLOAD_FOLDER'], 'data.xlsx', as_attachment=True)


@app.route("/checkup")
def checkup():
    # Проверяем статус сервера Росстата
    rosstat_status = check_rosstat_status()
    return render_template('checkup.html', rosstat_status=rosstat_status)


def check_rosstat_status():
    try:
        response = requests.get('https://showdata.gks.ru/finder/', verify=False)
        if response.status_code == 200:
            return 'Онлайн'
        else:
            return 'Оффлайн'
    except requests.RequestException:
        return 'Технические неполадки'


@app.route('/download_template')
def download_template():
    # Создаем шаблон DataFrame
    data = pd.DataFrame({
        'ds': pd.date_range(start='2005-01-01', periods=36, freq='M'),
        'y': [round(random.uniform(99.25, 102.54), 2) for _ in range(36)]
    })

    # Сохраняем DataFrame в файл Excel
    template_path = 'static/template.xlsx'
    data.to_excel(template_path, index=False)

    # Отправляем файл пользователю для скачивания
    return send_file(template_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)

