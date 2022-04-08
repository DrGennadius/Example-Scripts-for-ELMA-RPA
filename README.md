# Example-Scripts-for-ELMA-RPA
Примеры скриптов для [ELMA RPA](https://elma-rpa.ai/ru).
Некоторые примеры уже не так актуальны на новых версиях ELMA RPA. Но всё же пускай тут остаются, вдруг пригодятся.

ELMA RPA использует/генерирует проект (.csproj) со структурой:

| Наименование файла | Краткое описание |
| ---- | ---- |
| Script.csproj | Файл проекта. В текущих примерах заменено с другими названиями. |
| Context.cs | Файл с исходный кодом контекста. |
| Program.cs | Файл с исходный кодом консольной програмы для отлакдки скрипта без робота. |
| ScriptActivity.cs | Файл с исходный кодом скрипта. |
| Directory.Build.props | Ссылка на ELMA.RPA.SDK (<Путь установленного ELMA RPA Designer>/ELMA.RPA.SDK.dll) |

Пример проекта с пустым скриптом представлен в [Template](https://github.com/DrGennadius/Example-Scripts-for-ELMA-RPA/tree/master/Template).

Можно ознакомиться более детально с [описанием структуры временного проекта скрипта](https://github.com/DrGennadius/Example-Scripts-for-ELMA-RPA/wiki/%D0%9E%D0%BF%D0%B8%D1%81%D0%B0%D0%BD%D0%B8%D0%B5-%D1%81%D1%82%D1%80%D1%83%D0%BA%D1%82%D1%83%D1%80%D1%8B-%D0%B2%D1%80%D0%B5%D0%BC%D0%B5%D0%BD%D0%BD%D0%BE%D0%B3%D0%BE-%D0%BF%D1%80%D0%BE%D0%B5%D0%BA%D1%82%D0%B0-%D1%81%D0%BA%D1%80%D0%B8%D0%BF%D1%82%D0%B0).

Рекомендуется использовать Visual Studio 2019.

Проекты примеров находятся в папке "Примеры" в корне папки этого репозитория. Все примеры разбиты на категории (подпапки).

Если использовать Visual Studio, то достаточно открыть файл ".\Примеры\Примеры скриптов для ELMA RPA.sln". В Visual Studio все примеры будут также разбиты на категории:

![image](https://user-images.githubusercontent.com/27915885/162456938-61ec30ec-85b0-4790-812e-f6e333850087.png)

## Для тех, кто не умеет вот это всё, но очень хочется

Было бы не плохо, если уже есть какой-то опыт в программировании .NET (C#) и вас не пугают файлы с расширениями .cs, .csproj.

Если не так, то можно начать "поглядеть" в https://metanit.com/sharp/tutorial/ - в принципе излагается всё достаточно просто и без кучи лишнего.

А вообще по хорошому лучше бы тут https://docs.microsoft.com/ru-ru/dotnet/csharp/ - а тут может показаться сложно и запутанно, но зато более полноценно и точно.

Про проекты в Visual Studio 2019: https://docs.microsoft.com/ru-ru/visualstudio/ide/solutions-and-projects-in-visual-studio?view=vs-2019

## Visual Studio Code

Добавил еще файлы конфиграций сборки и запуска для отладки для Visual Studio Code. Чтобы не нужно было каждый раз генерировать. Необходимо открывать папку каждого примера (проекта, там где есть файл с расширением .csproj) в Visual Studio Code и сразу пользоваться без генерации.

## Дополнительно

Утилита для предотвращения сна Windows - [GSimpleWinSleepPreventer](https://github.com/DrGennadius/GSimpleWinSleepPreventer).
