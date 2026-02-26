"""
Script de teste para verificar se as notificações Windows estão funcionando
"""

print("🔔 Testando notificações Windows...")
print()

# Tentar win10toast primeiro
try:
    from win10toast import ToastNotifier
    print("✅ win10toast encontrado!")
    
    toaster = ToastNotifier()
    toaster.show_toast(
        "Teste de Notificação",
        "Se você está vendo isso, as notificações estão funcionando! 🎉",
        duration=10,
        threaded=True
    )
    print("✅ Notificação enviada com win10toast!")
    print("   Verifique se apareceu uma notificação no canto da tela.")
    
except ImportError:
    print("❌ win10toast não encontrado. Tentando plyer...")
    
    try:
        from plyer import notification
        print("✅ plyer encontrado!")
        
        notification.notify(
            title="Teste de Notificação",
            message="Se você está vendo isso, as notificações estão funcionando! 🎉",
            timeout=10,
            app_name="Monitor de Pedidos"
        )
        print("✅ Notificação enviada com plyer!")
        print("   Verifique se apareceu uma notificação no canto da tela.")
        
    except ImportError:
        print("❌ Nenhuma biblioteca de notificação encontrada!")
        print()
        print("📦 Para instalar, execute:")
        print("   pip install win10toast")
        print("   OU")
        print("   pip install plyer")
        print()
        print("💡 Recomendação: win10toast funciona melhor no Windows")

print()
print("=" * 60)
input("Pressione ENTER para sair...")



