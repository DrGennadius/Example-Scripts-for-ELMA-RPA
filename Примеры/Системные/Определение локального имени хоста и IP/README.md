Пригодится, если нужно определить с какой машины (какого бота) была выполнена задача.

Зачем может такое понадобиться? Вообще можно много придумать сценариев. Ниже перечислены некоторые примеры:

1. Когда бот на своей машине подготавливает какие-нибудь данные, которые нужно будет потом скачать с другой машины. Например, бот выполняет анализ, скачивает какие-нибудь данные, подготавливает/формирует "у себя" набор данных. Потом система, которая запустила задачу RPA, определяет с какой машины выполнялись операции и где сейчас подготовлены данные. Система скачивает подготовленные данные напрямую с той машины, где бот.
2. Для какой-нибудь сложной логики с распределённым выполнением. Сейчас, конечно, есть ограничения оркестратора. Задачи выполняются случайно на свободной системе (боте). Нельзя задать определенную систему (бота). Теоретически можно схитрить и дать доступ к файлам между системами и каждый бот будет знать и получать данные с других ботов. Т.е., например, мы имеем сложную логику и множество ботов, которые выполняют сложные операции и подготавливают много данных. И главное, что сложная логика определяется как процесс, который разделен на несколько этапы, которые выполняются как основной системой, так и местами RPA. Т.е. задачи RPA могут быть разные и выполняться в разное время и дополнять друг друга в рамках одного процесса.
3. Для статистики/мониторинга.
