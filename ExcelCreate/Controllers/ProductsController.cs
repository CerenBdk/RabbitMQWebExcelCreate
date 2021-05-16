using ExcelCreate.Models;
using ExcelCreate.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelCreate.Controllers
{
    [Authorize]
    public class ProductsController : Controller
    {

        private readonly UserManager<IdentityUser> _userManager;
        private readonly AppDbContext _context;
        private readonly RabbitMQPublisher _rabbitMQPublisher;
        public ProductsController(UserManager<IdentityUser> userManager, AppDbContext context,
            RabbitMQPublisher rabbitMQPublisher)
        {
            _userManager = userManager;
            _context = context;
            _rabbitMQPublisher = rabbitMQPublisher;
        }

        public IActionResult Index()
        {
            return View();
        }

        public async Task<IActionResult> CreateProductExcelAsync()
        {
            var user = await _userManager.FindByNameAsync(User.Identity.Name);

            var fileName = $"product-excel-{Guid.NewGuid().ToString().Substring(1, 20)}";
            UserFile userFile = new() { 
                UserId = user.Id,
                FileName = fileName,
                FileStatus = FileStatus.Creating
            };

            await _context.UserFiles.AddAsync(userFile);
            await _context.SaveChangesAsync();

            _rabbitMQPublisher.Publish(new() { FileId = userFile.Id});
            TempData["StartCreatingExcel"] = true;

            return RedirectToAction("Files");
        }

        public async Task<IActionResult> Files() 
        {
            var user = await _userManager.FindByNameAsync(User.Identity.Name);
            var result = await _context.UserFiles.Where(x => x.UserId == user.Id).OrderByDescending( x => x.Id).ToListAsync();

            return View(result);
        }
    }
}
